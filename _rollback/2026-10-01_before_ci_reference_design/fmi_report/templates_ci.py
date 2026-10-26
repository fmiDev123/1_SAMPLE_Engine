"""Competitive Landscape: ten competitive-intelligence pages (2026-09-30).

Modelled on the client's reference 'Aerial_Imaging_Competitive_Deck_10_pages'
(and its market worked example, the HER2 deck), built from this deck's own
blocks - KPI cards, native tables and summary panels - so they share its
colours, borders, 11pt Calibri and locking.  Every figure and rating comes
from `ci_data.build()` (see there for sources); a field only research can
supply reads 'n/a'.

    1  Competitive Landscape at a Glance
    2  Segment Leadership Map
    3  Value Chain Positioning of Leading Players
    4  Player Landscape by Business Type and Reach
    5  Corporate Scorecard: Who They Are
    6  Offering Scorecard: What They Sell
    7  Execution Scorecard: How They Win
    8  Success Factors and Composite Ranking
    9  Strategic Activity Tracker
    10 Strategic Activity by Function
"""
from __future__ import annotations

import re

import pandas as pd

from . import ci_data as X
from . import layout as L
from .components import fmt_num, fmt_pct
from .config import END_YEAR, EST_YEAR, UNIT_LABELS
from .data import MARKET_LABEL as MKT

SYM = {3: "●", 2: "◑", 1: "◔", 0: "○", None: "n/a"}
LEGEND = "● strong · ◑ moderate · ◔ weak · ○ none · n/a not assessed"
NA = "n/a"


def _short(name):
    """'Beverages (Water, ...)' -> 'Beverages' for a column heading."""
    return re.sub(r"\s*\([^)]*\)", "", str(name)).strip() or str(name)


def _cagr(a, b, n=END_YEAR - EST_YEAR):
    return (b / a) ** (1.0 / n) - 1.0 if a > 0 and b > 0 else float("nan")


def _co(c):
    """A company's row label: 'Apex Materials Corp (T1)'."""
    t = c["tier"].replace("Tier ", "T") if c["tier"] != "Others" else "Oth"
    return f"{c['name']} ({t})"


def _ci():
    return X.build()


def _top(ci, k):
    return sum(sorted((c["share"] for c in ci["companies"]),
                      reverse=True)[:k])


# ═══════════════════════ 1. At a glance ════════════════════════════════
def ci_glance():
    ci = _ci()
    u = UNIT_LABELS["Value"]
    cos = ci["companies"]
    prod, end = ci["prod_subs"], ci["end_subs"]
    mv0, mv1 = ci["mv0"], ci["mv1"]
    top3, top5 = _top(ci, 3), _top(ci, 5)
    p0 = prod[0] if prod else ("n/a", 0, 0)
    e0 = end[0] if end else ("n/a", 0, 0)
    kpis = [
        L.kpi(f"Top 3 share, {EST_YEAR}", f"{top3:.0f}%", "of market value",
              frac=1 / 6),
        L.kpi("Top 5 share", f"{top5:.0f}%", "of market value", frac=1 / 6),
        L.kpi(f"Largest {ci['prod_seg'] or 'segment'}".lower().capitalize(),
              fmt_pct(p0[1] / (mv0 or 1)), _short(p0[0]), frac=1 / 6),
        L.kpi("Largest end use", fmt_pct(e0[1] / (sum(r[1] for r in end)
                                                   or 1)),
              _short(e0[0]), frac=1 / 6),
        L.kpi("Named competitors", str(len(cos)), "profiled here",
              frac=1 / 6),
        L.kpi(f"Market CAGR {EST_YEAR}–{END_YEAR}",
              # no figure in a caption: a locked deck withholds only the
              # card's main value
              fmt_pct(_cagr(mv0, mv1)), "value, compound annual",
              frac=1 / 6),
    ]
    cols = [c for _k, c in ci["lead_cols"]]
    leaders = {c["name"] for c in cos if c["led"]}
    max_led = max((c["led"] for c in cos), default=0)
    fast = max(prod, key=lambda r: _cagr(r[1], r[2])) if prod else None
    t = ci["tiers"]
    n_moves = sum(1 for c in cos if len(c["active"]) >= 5)
    n_acq = sum(1 for c in cos if "A" in c["active"])
    rank = sorted(cos, key=lambda c: c["composite"], reverse=True)
    reach = {r: sum(1 for c in cos if c["reach"] == r)
             for r in ("Global", "Regional", "Local or niche")}
    structure = [
        f"The top three players hold {top3:.0f}% of {EST_YEAR} value; "
        f"{len(leaders)} companies lead the {len(cols)} sub-segments tracked "
        f"and no one leads more than {max_led}.",
        f"Tier 1 ({len(t['Tier 1']['players'])} players) holds "
        f"{fmt_pct(t['Tier 1']['share'])}; Tier 2 and Tier 3 together hold "
        f"{fmt_pct(t['Tier 2']['share'] + t['Tier 3']['share'])}.",
        f"{reach['Global']} players compete globally, {reach['Regional']} "
        f"regionally and {reach['Local or niche']} in one market or niche.",
    ]
    moves = [
        f"'{_short(p0[0])}' is the largest {str(ci['prod_seg']).lower()} "
        f"line ({fmt_pct(p0[1] / (mv0 or 1))})"
        + (f"; '{_short(fast[0])}' compounds fastest at "
           f"{fmt_pct(_cagr(fast[1], fast[2]))} CAGR." if fast else "."),
        f"'{_short(e0[0])}' is the largest end use; the top two end uses "
        f"take {fmt_pct(sum(r[1] for r in end[:2]) / (sum(r[1] for r in end) or 1))}"
        f" of value.",
        f"{n_moves} of {len(cos)} named players run five or more strategic "
        f"moves; acquisition is active at {n_acq}.  The composite ranking "
        f"puts {', '.join(c['name'] for c in rank[:3])} ahead.",
    ]
    method = [
        "Page 2 shows who leads which sub-segment; pages 3 and 4 show how "
        "the market is organised and who plays where.",
        "Pages 5 to 7 rate every named company on twenty Ps at corporate, "
        "offering and execution level; a cell reads n/a where no evidence "
        "exists.  Page 8 turns the pattern into five success factors.",
        "Pages 9 and 10 list each player's active strategic moves and what "
        "each department should do about them.",
    ]
    return (f"{MKT} Competitive Landscape at a Glance",
            {"rows": [L.frow(*kpis),
                      L.frow(L.summary("Market Structure", structure,
                                       frac=0.5, h=190),
                             L.summary("Where Competition Moves", moves,
                                       frac=0.5, h=190)),
                      L.frow(L.summary("How This Section Is Built", method,
                                       h=120))]})


# ═══════════════════════ 2. Segment leadership ═════════════════════════
def ci_leadership():
    ci = _ci()
    cos = ci["companies"]
    t = ci["tiers"]
    grp = pd.DataFrame({
        "Share of value": [fmt_pct(t[k]["share"]) for k in t],
        "Players": [", ".join(t[k]["players"]) or "Long tail of regional "
                    "and niche suppliers" for k in t]},
        index=list(t.keys()))
    cols = [c for _k, c in ci["lead_cols"]]
    heads = [_short(c) for c in cols]
    tbl = pd.DataFrame({h: [SYM[ci["leadership"][c["name"]][col]]
                            for c in cos] for h, col in zip(heads, cols)},
                       index=[_co(c) for c in cos])
    counts = {h: sum(1 for c in cos if ci["leadership"][c["name"]][col] >= 2)
              for h, col in zip(heads, cols)}
    most = max(counts, key=counts.get)
    least = min(counts, key=counts.get)
    read = [
        "Read across a row to see a player's breadth; down a column to see "
        "how contested a sub-segment is.",
        f"'{most}' is the most contested column ({counts[most]} players "
        f"active); '{least}' is the least ({counts[least]}), the more open "
        f"position.",
    ]
    return (f"{MKT} Segment Leadership Map", {"rows": [
        L.frow(L.tbl(grp, index_name="Player group", frac=0.64,
                     title=f"Who Holds the Market, {EST_YEAR}"),
               L.summary("Reading", read, frac=0.36, h=150)),
        L.frow(L.tbl(tbl, index_name="Company (tier)",
                     title=f"Who Leads Which Sub-segment, {EST_YEAR}  "
                           f"(● leader · ◑ active · ○ not active)")),
    ]})


# ═══════════════════════ 3. Value chain positioning ════════════════════
def ci_value_chain():
    ci = _ci()
    cos = ci["companies"]
    prod, end = ci["prod_subs"], ci["end_subs"]
    mv0 = ci["mv0"] or 1
    lead = ci["leadership"]

    def who(col, lvl):
        return [c["name"] for c in cos if lead[c["name"]].get(col) == lvl]

    panels = []
    for i, (name, v0, v1) in enumerate(prod[:3], start=1):
        leads = who(name, 3) or ["n/a"]
        also = who(name, 2)[:3] or ["n/a"]
        panels.append(L.summary(
            f"{i}. {_short(name)}",
            [f"{fmt_pct(v0 / mv0)} of market, {fmt_pct(_cagr(v0, v1))} CAGR",
             "Leads: " + ", ".join(leads),
             "Also active: " + ", ".join(also)], frac=0.25, h=210))
    etot = sum(r[1] for r in end) or 1
    panels.append(L.summary(
        f"{len(prod[:3]) + 1}. End use",
        [f"{_short(n)}: {fmt_pct(v / etot)}" for n, v, _w in end[:5]],
        frac=0.25, h=210))
    try:
        from . import dynamics as D
        enab = [d["title"] for d in D.items("drivers")][:3]
        cons = [d["title"] for d in D.items("restraints")][:3]
    except Exception:
        enab, cons = [NA], [NA]
    kpis = [L.kpi(f"Top 3 share, {EST_YEAR}", f"{_top(ci, 3):.0f}%", "",
                  frac=0.2),
            L.kpi("Top 5 share", f"{_top(ci, 5):.0f}%", "", frac=0.2)] + \
        [L.kpi(_short(n), fmt_pct(v / mv0), "of market", frac=0.2)
         for n, v, _w in prod[:3]]
    return (f"{MKT} Value Chain Positioning of Leading Players", {"rows": [
        L.frow(*kpis),
        L.frow(*panels),
        L.frow(L.summary("Enablers", enab, frac=0.5, h=130),
               L.summary("Constraints", cons, frac=0.5, h=130)),
    ]})


# ═══════════════════════ 4. Player landscape ═══════════════════════════
def ci_players():
    ci = _ci()
    cos = ci["companies"]
    types = list(dict.fromkeys((c["type"] or c["tier"]) for c in cos))
    reaches = ["Global", "Regional", "Local or niche"]
    grid = pd.DataFrame(
        {tp: ["; ".join(f"{c['name']}" + (f" ({c['hq']})" if c["hq"] else "")
                        for c in cos if (c["type"] or c["tier"]) == tp
                        and c["reach"] == r) or "–" for r in reaches]
         for tp in types},
        index=[f"{r} ({sum(1 for c in cos if c['reach'] == r)})"
               for r in reaches])
    panels = []
    for tp in types[:4]:
        mem = [c for c in cos if (c["type"] or c["tier"]) == tp]
        dom = max(set(c["purpose"] for c in mem),
                  key=lambda p: sum(1 for c in mem if c["purpose"] == p))
        panels.append(L.summary(tp, [
            f"{len(mem)} player{'s' if len(mem) != 1 else ''}, "
            f"{fmt_pct(sum(c['share'] for c in mem) / 100)} combined share",
            f"Route to market: {mem[0]['route']}",
            f"Dominant strategy: {dom}"], frac=1 / min(4, len(types)),
            h=150))
    return (f"{MKT} Player Landscape by Business Type and Reach", {"rows": [
        L.frow(L.tbl(grid, index_name="Reach",
                     title=f"Named Players by Business Type and Reach, "
                           f"{EST_YEAR}")),
        L.frow(*panels),
    ]})


# ═══════════════════════ 5-7. Scorecards ═══════════════════════════════
def _scorecard(title, cols, note):
    ci = _ci()
    cos = ci["companies"]
    tbl = pd.DataFrame({h: [f(c) for c in cos] for h, f in cols},
                       index=[_co(c) for c in cos])
    return (f"{MKT} {title}", {"rows": [
        L.frow(L.tbl(tbl, index_name="Company (tier)",
                     title=f"{title.split(':')[0]}, {EST_YEAR}  ({LEGEND})")),
        L.frow(L.summary("How the Ratings Are Set", note, h=120)),
    ]})


def ci_corporate():
    return _scorecard("Corporate Scorecard: Who They Are", [
        ("Parent and ownership", lambda c: c["type"] or NA),
        ("Purpose (strategy)", lambda c: c["purpose"]),
        ("Portfolio breadth", lambda c: SYM[c["s_portfolio"]]),
        ("Presence", lambda c: SYM[c["s_presence"]]),
        ("Power (scale)", lambda c: SYM[c["s_power"]]),
        ("Performance", lambda c: SYM[c["s_performance"]]),
        ("People", lambda c: NA),
        ("Partnerships", lambda c: NA),
    ], [
        "Portfolio breadth: sub-segments led or active in, of eight (6+ "
        "strong, 4-5 moderate, 2-3 weak).  Presence: global 3, regional 2, local or niche 1.",
        "Power: Tier 1 = 3, Tier 2 = 2, Tier 3 = 1.  Performance: share of "
        "market value (10%+ strong, 5-10% moderate, below 5% weak).",
        "People and Partnerships need company-level research and read n/a "
        "until the research feed supplies them.",
    ])


def ci_offering():
    ci = _ci()
    patents = "In full study" if X.covered(["patent"], ci) else NA
    return _scorecard("Offering Scorecard: What They Sell", [
        ("Product", lambda c: SYM[c["s_product"]]),
        ("Platform (technology)", lambda c: SYM[c["s_platform"]]),
        ("Price and model", lambda c: c["price"]),
        ("Positioning (verticals)",
         lambda c: ", ".join(_short(v) for v in c["verticals"][:2]) or NA),
        ("Proof", lambda c: NA),
        ("Patents and IP", lambda c: patents),
    ], [
        "Product: product types offered of the market's top three.  "
        "Platform: R&D investment and product development both in motion "
        "= strong, one = moderate.",
        "Price and model follows the tier's positioning; Positioning names "
        "the end uses a player serves.",
        "Proof (named customers) and Patents need research; the full study "
        "covers patents where its contents list them.",
    ])


def ci_execution():
    return _scorecard("Execution Scorecard: How They Win", [
        ("Place (channel)", lambda c: c["route"]),
        ("Promotion", lambda c: NA),
        ("Pace (momentum)", lambda c: SYM[c["s_pace"]]),
        ("Pipeline", lambda c: c["pipeline"] or NA),
        ("Policy exposure", lambda c: NA),
        ("Prospects", lambda c: SYM[c["s_prospects"]]),
    ], [
        "Pace: active strategic moves (5+ strong, 3-4 moderate, fewer weak).  "
        "Pipeline names a launch or development programme in motion.",
        "Prospects: analyst view from breadth, reach, scale and pace (strong "
        "= gaining, moderate = holding, weak = at risk).",
        "Promotion and Policy exposure need market and regulatory research "
        "and read n/a until supplied.",
    ])


# ═══════════════════════ 8. Success factors ════════════════════════════
def ci_success():
    ci = _ci()
    cos = sorted(ci["companies"], key=lambda c: c["composite"], reverse=True)
    f = ci["factors"]
    tbl = pd.DataFrame(
        {f"{i}. {n} ({int(w * 100)}%)": [SYM[c[k]] for c in cos]
         for i, (n, w, k, _d) in enumerate(f, start=1)},
        index=[_co(c) for c in cos])
    tbl["Composite (0–3)"] = [f"{c['composite']:.2f}" for c in cos]
    groups = {}
    for c in cos:
        groups.setdefault(c["type"] or c["tier"], []).append(c)
    panels = []
    for tp, mem in list(groups.items())[:4]:
        best = max(f, key=lambda x: sum(c[x[2]] for c in mem))
        worst = min(f, key=lambda x: sum(c[x[2]] for c in mem))
        panels.append(L.summary(tp, [
            ", ".join(c["name"] for c in mem),
            f"Wins on {best[0].lower()}; weakest on {worst[0].lower()}."],
            frac=1 / min(4, len(groups)), h=140))
    return (f"{MKT} Success Factors and Composite Ranking", {"rows": [
        L.frow(L.tbl(tbl, index_name="Company (tier)",
                     title=f"Five Success Factors, Weighted, {EST_YEAR}  "
                           f"({LEGEND})")),
        L.frow(*panels),
    ]})


# ═══════════════════════ 9. Activity tracker ═══════════════════════════
def _tracked(ci, limit=12):
    """Up to `limit` moves, the most telling first: acquisitions, launches,
    development, expansion - one pass per company before a second."""
    order = {"MA": 0, "PR": 1, "RD": 2, "EX": 3, "RG": 4, "CP": 5}
    # one move per player: the most telling one
    by_co = {}
    for a in sorted(ci["activity"], key=lambda a: order.get(a["code"], 9)):
        by_co.setdefault(a["company"], []).append(a)
    out, k = [], 0
    while len(out) < limit and any(len(v) > k for v in by_co.values()):
        for co in [c["name"] for c in ci["companies"]]:
            if len(by_co.get(co, [])) > k and len(out) < limit:
                out.append(by_co[co][k])
        k += 1
    return out


def ci_activity():
    ci = _ci()
    rows = _tracked(ci, limit=len(ci["companies"]))
    tbl = pd.DataFrame({
        "Development (what the player is doing)": [a["what"] for a in rows],
        "Type": [X.TYPE_NAMES.get(a["code"], a["code"]) for a in rows],
        "Source department": [a["dept"] for a in rows],
        "Signal (what it indicates)": [a["signal"] for a in rows],
        "Who should act": [a["act"] for a in rows],
    }, index=[a["company"] for a in rows])
    return (f"{MKT} Strategic Activity Tracker", {"rows": [
        L.frow(L.tbl(tbl, index_name="Company",
                     title=f"Most Telling Active Move by Player, {EST_YEAR}  "
                           f"(dates, deal values and evidence populate from "
                           f"the research feed)")),
    ]})


# ═══════════════════════ 10. Activity by function ══════════════════════
def ci_functions():
    ci = _ci()
    rows = ci["activity"]
    cos = [c["name"] for c in ci["companies"]]
    codes = [k for k in X.TYPE_NAMES if any(a["code"] == k for a in rows)]
    by_type = pd.DataFrame(
        {X.TYPE_NAMES[k]: [str(sum(1 for a in rows if a["company"] == co
                                   and a["code"] == k) or "–") for co in cos]
         for k in codes}, index=cos)
    depts = [d for d in X.DEPARTMENTS if any(a["dept"] == d for a in rows)]
    by_dept = pd.DataFrame(
        {d: [str(sum(1 for a in rows if a["company"] == co and a["dept"] == d)
                 or "–") for co in cos] for d in depts}, index=cos)

    def n_of(code):
        return sum(1 for a in rows if a["code"] == code)
    meaning = [
        ("Sales", f"{n_of('EX')} expansion moves target new regions and "
                  f"volume tiers: defend key accounts where rivals expand."),
        ("Product and R&D", f"{n_of('PR') + n_of('RD')} product and R&D "
                            f"moves: benchmark launches and the technology "
                            f"roadmap."),
        ("Corporate development", f"{n_of('MA')} players are acquiring: the "
                                  f"independent target list is shrinking."),
        ("Legal and regulatory", f"{n_of('RG')} players use certification "
                                 f"as a trust signal: audit your own."),
    ]
    return (f"{MKT} Strategic Activity by Function", {"rows": [
        L.frow(L.tbl(by_type, index_name="Company", frac=0.5,
                     title="Moves by Type (count)"),
               L.tbl(by_dept, index_name="Company", frac=0.5,
                     title="Moves by Source Department (count)")),
        L.frow(*[L.summary(h, [t], frac=0.25, h=130) for h, t in meaning]),
    ]})


PAGES = (
    ("Competitive Landscape at a Glance", ci_glance),
    ("Segment Leadership Map", ci_leadership),
    ("Value Chain Positioning", ci_value_chain),
    ("Player Landscape", ci_players),
    ("Corporate Scorecard", ci_corporate),
    ("Offering Scorecard", ci_offering),
    ("Execution Scorecard", ci_execution),
    ("Success Factors", ci_success),
    ("Strategic Activity Tracker", ci_activity),
    ("Strategic Activity by Function", ci_functions),
)
