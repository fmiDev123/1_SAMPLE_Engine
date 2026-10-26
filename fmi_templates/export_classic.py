"""
Build FMI_Classic_Report_Templates.xlsx (2015 report style) from classic.py.

    python -m fmi_templates.export_classic [output.xlsx]
"""
import sys

from openpyxl import Workbook
from openpyxl.styles import Font

from .classic import CLASSIC, CLASSIC_DYNAMICS
from .slide_templates import SLIDES
from .demo_classic import dwm_check, muffin_render
from .export_workbook import _sheet, _flatten, SECTION_FILL, BOLD, BODY_FONT, WRAP, NAVY


def build(out):
    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    rows = [
        ("FMI Classic Report Templates (2015 style)", None),
        ("What this is", "Slide content written in the voice and layout of FMI's 2015 syndicated reports (reference: REP-GB-717 Digital Wound Measurement Devices, 2015-2025). Every data slide has a title, an italic headline with the key number, auto-numbered Figure and Table captions, and a Market Outlook box. Only numbers, segment and region names and the figure and table headings change between reports."),
        ("Proof", "The '2015 Check' sheet regenerates nine pages of the 2015 report from its own numbers and puts the original text beside the generated text."),
        ("Numbering", "render_deck(plan) assigns Figure 01, 02 ... and Table 01, 02 ... in deck order, fills 'Figures 09 and 10 highlight ...' automatically, and returns the List of Figures and List of Tables pages."),
        ("Rules kept from the 2015 voice", "'projected to register 1.7X increase', 'X segment is projected to register 2.1X higher value than Y segment in 2025', 'gaining / losing n BPS', 'Absolute $ value for 2026 is estimated to be ...'. Direction words (increase, decline, stable; gaining, losing) are picked from the numbers, so a falling segment never reads as an increase."),
        ("Changes from the 2015 original", "Hyphens replace en dashes in year ranges (house style). Bullets end with full stops. Known slips in the 2015 report are not carried over: Figures 17 and 18 captioned 'Incisional' on the traumatic-wounds page, 'Hospitals end-user segment in Japan' on the Western Europe page, 'By Region' on single-region figures, and the North America incremental opportunity quoted as the 2025 value."),
        ("Placeholders", "Same convention as the modern library: v_ market figure (XX in samples; multiples show X.X), c_ company data, n_ count, s_ score, in_ research input, cp_ classic rule phrase. Sentence-start and mid-sentence forms of names are built automatically (item_lc, item_title, mkt_short_lc, geo_the)."),
        ("Qualitative slides", "Drivers, restraints and trends default to the industry set on the 'Dynamics by Industry' sheet (food, consumer, healthcare, pharma, industrial). Replace with report-specific inputs where available. Statistics ('For instance, ...') must come from research input; the templates never invent them."),
        ("Sheets", "Classic Templates | Captions | Current Deck Mapping | Dynamics by Industry | 2015 Check | Rendered Muffins | Muffins LoF LoT"),
    ]
    for i, (k, v) in enumerate(rows, 1):
        ws.cell(row=i, column=1, value=k).font = Font(name="Arial", bold=True, size=14 if i == 1 else 10, color=NAVY if i == 1 else "000000")
        if v:
            c = ws.cell(row=i, column=2, value=v)
            c.font, c.alignment = BODY_FONT, WRAP
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 120

    trows, crow = [], []
    for t in CLASSIC:
        base = (t["id"], t["chapter"], t["name"], t["ref"], t["maps_to"])
        trows.append(base + ("title", 1, t["title"]))
        if t["headline"]:
            trows.append(base + ("headline (italic)", 1, t["headline"]))
        for i, f in enumerate(t["figures"], 1):
            trows.append(base + ("figure caption", i, f"Figure {{nn}}: {f}"))
            crow.append((t["id"], t["name"], "Figure", f))
        for i, f in enumerate(t["tables"], 1):
            trows.append(base + ("table caption", i, f"Table {{nn}}: {f}"))
            crow.append((t["id"], t["name"], "Table", f))
        for key, val in t["elements"].items():
            label = key.replace("__", " (") + (")" if "__" in key else "")
            for i, text, _ in _flatten(val):
                trows.append(base + (label, i, text))
        trows.append(base + ("source", 1, "Source: FMI, {pub_year}"))
    wst = _sheet(wb, "Classic Templates", ["ID", "Chapter", "Slide type", "2015 report pages", "Replaces current IDs", "Element", "#", "Template text"],
                 trows, [7, 16, 30, 12, 14, 18, 5, 110], "C2")
    last = None
    for r in range(2, wst.max_row + 1):
        tid = wst.cell(row=r, column=1).value
        if tid != last:
            for c in range(1, 9):
                wst.cell(row=r, column=c).fill = SECTION_FILL
            wst.cell(row=r, column=8).font = BOLD
            last = tid

    _sheet(wb, "Captions", ["ID", "Slide type", "Kind", "Caption pattern (number added automatically)"], crow, [7, 34, 8, 110])

    cmap = {}
    for t in CLASSIC:
        for m in [x.strip() for x in t["maps_to"].split(",") if x.strip()]:
            cmap.setdefault(m, []).append(t["id"])
    mrows = [(s["id"], s["name"], ", ".join(cmap.get(s["id"], [])) or "Keep modern template (no 2015 equivalent)") for s in SLIDES]
    _sheet(wb, "Current Deck Mapping", ["Current template ID", "Current slide type", "Classic template(s) to use"], mrows, [14, 44, 50])

    drows = []
    for ind, d in CLASSIC_DYNAMICS.items():
        drows.append((ind, "Drivers headline", "", d["drivers_headline"]))
        drows.append((ind, "Restraints headline", "", d["restraints_headline"]))
        for kind in ("drivers", "restraints", "trends"):
            for title, text in d[kind]:
                drows.append((ind, kind[:-1].capitalize(), title, text))
    _sheet(wb, "Dynamics by Industry", ["Industry", "Type", "Title", "Text"], drows, [12, 18, 40, 110])

    chk = [(cid, name, orig, gen, str(miss or "")) for cid, name, orig, gen, miss in dwm_check()]
    _sheet(wb, "2015 Check", ["ID", "Slide type", "Original 2015 report text (REP-GB-717)", "Generated from the template with the same numbers", "Missing"],
           chk, [7, 26, 75, 75, 10], "C2")

    slides, lof, lot, issues = muffin_render()
    rrows = []
    for s in slides:
        for k, v in s.items():
            if k in ("id", "name", "footer", "_missing"):
                continue
            for i, text, _ in _flatten(v):
                rrows.append((s["id"], s["name"], k, i, text))
    _sheet(wb, "Rendered Muffins", ["ID", "Slide type", "Element", "#", "Rendered text (Packaged Muffins, masked sample)"], rrows, [7, 30, 18, 5, 120], "C2")
    _sheet(wb, "Muffins LoF LoT", ["List", "Entry"], [("List of Figures", x) for x in lof] + [("List of Tables", x) for x in lot], [16, 110])
    wb.save(out)
    return out, len(issues)


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "FMI_Classic_Report_Templates.xlsx"))
