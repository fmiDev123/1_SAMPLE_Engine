"""Table of Contents, List of Figures and List of Tables.

The research team writes the report's table of contents in Word.  Until now it
was pasted into the deck by hand - thirty rows at a time, two tables to a
slide, with "Keep Source Formatting" so Word's own numbering, indent and
colour survived the trip.  This module supplies the content for that, so the
engine can lay the same pages out on its own.

Two jobs live here:

  * `parse_toc` reads the Word file.  The numbers in front of each line
    ("2.", "2.1.", "2.2.1.1.") are NOT text in that file - they are Word's
    automatic list numbering, worked out at display time from
    `numbering.xml`, and for nearly every one of these paragraphs reached
    only through the paragraph STYLE ("List Number 3" -> numId 1, level 2)
    rather than through the paragraph itself.  So Word's numbering is
    re-implemented here, counters and all; read the paragraph text alone and
    every line arrives bare.

  * `figure_entries` / `table_entries` turn what the deck actually drew into
    the List of Figures and the List of Tables, numbered in slide order.

`export_native` owns the drawing; this module owns the content.
"""
from __future__ import annotations

import os
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

_PROJ_ROOT = Path(__file__).resolve().parents[1]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ── the file ─────────────────────────────────────────────────────────

def find_toc_file(base: Path | None = None) -> Path | None:
    """Locate the Word table of contents, or None if there isn't one.

    Priority: the REPORT_TOC env var, then 'TOC_*.docx' - the same shape of
    convention `config._find_data_file` uses for the workbook, so a market is
    set up by dropping two files in one folder - then any .docx whose name
    mentions a TOC.  Returns None rather than raising: a deck without a TOC
    file is built without the TOC pages, not failed.
    """
    env = os.environ.get("REPORT_TOC")
    if env:
        p = Path(env)
        if p.exists():
            return p
        cand = _PROJ_ROOT / env
        if cand.exists():
            return cand
        # Named a TOC that is not there.  Say so: falling through to the glob
        # would quietly paste a DIFFERENT market's contents into the deck.
        print(f"WARNING: REPORT_TOC={env!r} does not exist; falling back to "
              f"whatever TOC is in {_PROJ_ROOT}", file=sys.stderr)

    roots = [r for r in (base, _PROJ_ROOT, _PROJ_ROOT.parent, Path.cwd())
             if r is not None]
    # A loose "*toc*.docx" is only trusted beside the workbook being built:
    # picked up from the project root it pastes ANOTHER market's contents
    # into the deck (a polybag TOC in a hair-styling report).
    data_dir = None
    try:
        from . import data as _data
        data_dir = Path(str(_data.DATA_FILE)).resolve().parent
    except Exception:
        pass
    for pattern in ("TOC_*.docx", "TOC_*.DOCX", "*.docx"):
        for root in roots:
            if pattern.startswith("*") and data_dir is not None                     and Path(root).resolve() != data_dir:
                continue
            hits = [p for p in sorted(Path(root).glob(pattern))
                    if not p.name.startswith("~$")          # Word lock files
                    and (not pattern.startswith("*")
                         or "toc" in p.name.lower())]
            if hits:
                return hits[0]
    return None


# ── one line of the contents ─────────────────────────────────────────

@dataclass
class TocEntry:
    """One row of the contents table, already resolved to what it looks like.

    `marker` is the rendered list number or bullet ("2.1.", "o", "") and
    `marker_font` the face it has to be set in - a Symbol bullet is a
    private-use codepoint and turns into a box in any other font.  `marL` is
    the cell's left inset in EMU, which is how the indent is carried: these
    are single-column tables, so there is no hanging indent to hang from.
    """
    text: str
    level: int = 0
    marker: str = ""
    marker_font: str = "Calibri"
    size: float = 10.0                    # points
    bold: bool = False
    italic: bool = False
    color: str = "000000"                 # RRGGBB, no '#'
    marL: int = 9525                      # EMU
    hanging: float = 18.0                 # points - see `tab_pad`
    bullet: bool = False

    def tab_pad(self) -> int:
        """How many spaces stand in for the tab between number and text.

        Word puts a tab there and lets the paragraph's hanging indent decide
        where the text lands.  A table cell has no hanging indent to hang
        from, so the gap has to be padded out by hand - which is exactly what
        PowerPoint did when the team pasted this in, and matching its count
        is what makes the two decks line up.  Pad to the hanging indent, and
        every line of a level starts in the same column, which the hand-made
        deck does not quite manage ("23.10." came out with a gaping tab).
        """
        if not self.marker:
            return 0
        # A bullet is ONE glyph of a symbol face, measured as a whole; a
        # number is digits and dots in Calibri, measured character by
        # character.  An unfamiliar bullet font still takes the symbol path,
        # or its private-use glyph would be measured as if it were a digit.
        em = _MARKER_EM.get(self.marker_font) or (0.35 if self.bullet else 0)
        width = (em * len(self.marker) if em
                 else sum(_CALIBRI_EM.get(c, 0.507) for c in self.marker))
        return max(0, round((self.hanging - width * self.size) / _PAD_PT))


# One space of Times New Roman 7pt - the run PowerPoint pads a tab with.
_PAD_FONT, _PAD_SIZE = "Times New Roman", 7.0
_PAD_PT = _PAD_SIZE * 0.25
# Advance widths, in ems, for the glyphs a list marker is made of.
_CALIBRI_EM = {".": 0.267, "(": 0.303, ")": 0.303, " ": 0.226}
_MARKER_EM = {"Symbol": 0.35, "Courier New": 0.6, "Wingdings": 0.35}


# The ladder below is not invented - it is measured off the deck the research
# team produced by hand ("toc lot lof polybag market.pptx"), counting how each
# level actually came out of a Keep-Source-Formatting paste.  Levels 0-3 are
# what real files use; 4 continues the spacing so a deeper TOC still indents.
_LEVEL_STYLE = {
    0: dict(marL=171450,  size=10.0, bold=True,  italic=False, color="000000"),
    1: dict(marL=428625,  size=10.0, bold=False, italic=False, color="0070C0"),
    2: dict(marL=771525,  size=9.0,  bold=False, italic=False, color="000000"),
    3: dict(marL=1028700, size=9.0,  bold=False, italic=True,  color="000000"),
    4: dict(marL=1257300, size=9.0,  bold=False, italic=True,  color="000000"),
}
# A bulleted level sits further right than a numbered one at the same depth,
# because the bullet glyph is narrower than "2.1." and Word pads to a stop.
_BULLET_MARL = {0: 171450, 1: 685800, 2: 1028700, 3: 1257300, 4: 1485900}


def _level_style(level: int, bullet: bool) -> dict:
    st = dict(_LEVEL_STYLE[min(level, max(_LEVEL_STYLE))])
    if bullet:
        st["marL"] = _BULLET_MARL.get(level, st["marL"])
    return st


# ── Word's list numbering, re-implemented ────────────────────────────

_ROMAN = ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
          (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
          (5, "v"), (4, "iv"), (1, "i"))


def _roman(n: int) -> str:
    out = []
    for value, sign in _ROMAN:
        while n >= value:
            out.append(sign)
            n -= value
    return "".join(out)


def _letter(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("a") + rem) + out
    return out


def _format_counter(n: int, fmt: str) -> str:
    if fmt == "decimal":
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "lowerLetter":
        return _letter(n)
    if fmt == "upperLetter":
        return _letter(n).upper()
    if fmt == "lowerRoman":
        return _roman(n)
    if fmt == "upperRoman":
        return _roman(n).upper()
    if fmt == "none":
        return ""
    return str(n)


def _val(node, tag):
    el = node.find(_W + tag)
    return el.get(_W + "val") if el is not None else None


class _Numbering:
    """`numbering.xml` + `styles.xml`, far enough to number a paragraph.

    Word keeps a list definition in two halves: `w:num` (the instance a
    paragraph points at) resolves to a `w:abstractNum` (the shape of the
    list - how each level is formatted).  A paragraph reaches a `w:num`
    either directly, through its own `w:numPr`, or - as almost every line of
    these contents does - through its paragraph style.
    """

    def __init__(self, zf: zipfile.ZipFile):
        self._abstract: dict[str, dict[int, dict]] = {}
        self._num_to_abstract: dict[str, str] = {}
        self._style_num: dict[str, tuple[str | None, int | None]] = {}
        self._style_base: dict[str, str] = {}
        self._load_numbering(zf)
        self._load_styles(zf)
        # counters, per abstract list: level -> value, 0 meaning "not started"
        self._counters: dict[str, list[int]] = {}

    # -- parsing ------------------------------------------------------
    def _load_numbering(self, zf):
        try:
            root = etree.fromstring(zf.read("word/numbering.xml"))
        except KeyError:
            return
        for an in root.iter(_W + "abstractNum"):
            aid = an.get(_W + "abstractNumId")
            levels = {}
            for lvl in an.findall(_W + "lvl"):
                try:
                    ilvl = int(lvl.get(_W + "ilvl"))
                except (TypeError, ValueError):
                    continue
                fonts = lvl.find(_W + "rPr/" + _W + "rFonts")
                ind = lvl.find(_W + "pPr/" + _W + "ind")
                hang = ind.get(_W + "hanging") if ind is not None else None
                levels[ilvl] = {
                    "start": int(_val(lvl, "start") or 1),
                    "numFmt": _val(lvl, "numFmt") or "decimal",
                    "lvlText": _val(lvl, "lvlText") or "",
                    "restart": _val(lvl, "lvlRestart"),
                    "font": (fonts.get(_W + "ascii") if fonts is not None
                             else None),
                    # twips -> points; the tab after the number runs to here
                    "hanging": (int(hang) / 20.0 if hang and hang.isdigit()
                                else 18.0),
                }
            self._abstract[aid] = levels
        for num in root.iter(_W + "num"):
            nid = num.get(_W + "numId")
            aid = num.find(_W + "abstractNumId")
            if nid is not None and aid is not None:
                self._num_to_abstract[nid] = aid.get(_W + "val")

    def _load_styles(self, zf):
        try:
            root = etree.fromstring(zf.read("word/styles.xml"))
        except KeyError:
            return
        for st in root.iter(_W + "style"):
            sid = st.get(_W + "styleId")
            if not sid:
                continue
            based = st.find(_W + "basedOn")
            if based is not None:
                self._style_base[sid] = based.get(_W + "val")
            pPr = st.find(_W + "pPr")
            numPr = pPr.find(_W + "numPr") if pPr is not None else None
            if numPr is None:
                continue
            nid = numPr.find(_W + "numId")
            ilvl = numPr.find(_W + "ilvl")
            self._style_num[sid] = (
                nid.get(_W + "val") if nid is not None else None,
                int(ilvl.get(_W + "val")) if ilvl is not None else None,
            )

    def _style_numbering(self, style_id):
        """(numId, ilvl) for a style, following basedOn for whichever half the
        style itself leaves out - 'List Number 5' sets only the level and
        inherits the list from 'List Number 4'."""
        num_id = ilvl = None
        seen = set()
        sid = style_id
        while sid and sid not in seen:
            seen.add(sid)
            got = self._style_num.get(sid)
            if got:
                if num_id is None:
                    num_id = got[0]
                if ilvl is None:
                    ilvl = got[1]
            if num_id is not None and ilvl is not None:
                break
            sid = self._style_base.get(sid)
        return num_id, ilvl

    # -- use ----------------------------------------------------------
    def resolve(self, p, style_id):
        """(numId, ilvl) in force for this paragraph, or (None, ilvl).

        A direct `w:numPr` wins over the style's, and numId "0" is Word for
        "numbering switched off here" - the paragraph keeps its style's
        indent level but gets no number.
        """
        num_id = ilvl = None
        pPr = p.find(_W + "pPr")
        numPr = pPr.find(_W + "numPr") if pPr is not None else None
        if numPr is not None:
            nid = numPr.find(_W + "numId")
            il = numPr.find(_W + "ilvl")
            if nid is not None:
                num_id = nid.get(_W + "val")
            if il is not None:
                ilvl = int(il.get(_W + "val"))
        s_num, s_ilvl = self._style_numbering(style_id)
        if num_id is None:
            num_id = s_num
        if ilvl is None:
            ilvl = s_ilvl if s_ilvl is not None else 0
        if num_id in (None, "0"):
            return None, ilvl
        return num_id, ilvl

    def advance(self, num_id: str, ilvl: int):
        """Take the next number at `ilvl`, and render it.

        Returns (marker, font, is_bullet).  The counter rule is Word's, and
        the part of it that is easy to get wrong is this: a paragraph at a
        deep level COMMITS every shallower counter that has not started yet
        to its start value.  That is why a contents file whose first numbered
        line is a level-4 item ("1.1.1.1. Key Supply Trends") makes the next
        level-1 heading "2." and not "1." - the manual paste shows exactly
        that, and so must this.
        """
        aid = self._num_to_abstract.get(num_id)
        levels = self._abstract.get(aid) or {}
        lvl = levels.get(ilvl)
        if lvl is None:
            return "", "Calibri", False, 18.0
        counters = self._counters.setdefault(aid, [0] * 9)
        for j in range(min(ilvl, 8)):
            if counters[j] == 0:
                counters[j] = (levels.get(j) or {}).get("start", 1)
        if ilvl < 9:
            start = lvl.get("start", 1)
            counters[ilvl] = start if counters[ilvl] == 0 else counters[ilvl] + 1
            for j in range(ilvl + 1, 9):
                counters[j] = 0

        text = lvl["lvlText"]
        hang = lvl.get("hanging", 18.0)
        if lvl["numFmt"] == "bullet":
            return text, lvl.get("font") or "Symbol", True, hang

        def sub(m):
            j = int(m.group(1)) - 1
            if not 0 <= j < 9:
                return m.group(0)
            n = counters[j] or (levels.get(j) or {}).get("start", 1)
            return _format_counter(n, (levels.get(j) or {}).get("numFmt",
                                                               "decimal"))

        return (re.sub(r"%(\d)", sub, text),
                lvl.get("font") or "Calibri", False, hang)


# ── reading the document ─────────────────────────────────────────────

def _run_format(p) -> dict:
    """Bold / italic / colour / size off the paragraph's first real run.

    Only used for the lines that carry no number - the section banners and
    the red "Note:" lines the research team writes in by hand.  Those keep
    whatever they look like in Word rather than taking a level's styling.
    """
    out = {}
    for r in p.findall(_W + "r"):
        if not "".join(r.itertext()).strip():
            continue
        rPr = r.find(_W + "rPr")
        if rPr is None:
            break
        if rPr.find(_W + "b") is not None:
            out["bold"] = (rPr.find(_W + "b").get(_W + "val")
                           not in ("0", "false"))
        if rPr.find(_W + "i") is not None:
            out["italic"] = (rPr.find(_W + "i").get(_W + "val")
                             not in ("0", "false"))
        col = rPr.find(_W + "color")
        if col is not None:
            v = (col.get(_W + "val") or "").upper()
            if re.fullmatch(r"[0-9A-F]{6}", v) and v != "AUTO":
                out["color"] = v
        sz = rPr.find(_W + "sz")
        if sz is not None:
            try:
                out["size"] = int(sz.get(_W + "val")) / 2.0
            except (TypeError, ValueError):
                pass
        break
    return out


def _paragraph_text(p) -> str:
    """The paragraph's visible text, tabs flattened, deleted text dropped."""
    parts = []
    for node in p.iter():
        tag = node.tag
        if tag == _W + "t":
            parts.append(node.text or "")
        elif tag == _W + "tab":
            parts.append(" ")
        elif tag in (_W + "br", _W + "cr"):
            parts.append(" ")
    return re.sub(r"\s+", " ", "".join(parts)).strip()


def parse_toc(path: str | Path) -> list[TocEntry]:
    """Read the Word contents into the rows the deck will show.

    The contents are the stretch of the document from its first numbered
    "List Number" paragraph to its last.  Bounding it that way is what keeps
    the company address block out: those paragraphs carry the List Number
    style too, but with numbering switched off, and they sit after the final
    numbered line ("25. Research Methodology").  Unnumbered lines BETWEEN two
    numbered ones are kept - that is where the section banners and the
    "Note:" lines live, and the manual deck shows all of them.
    """
    path = Path(path)
    with zipfile.ZipFile(path) as zf:
        numbering = _Numbering(zf)
        root = etree.fromstring(zf.read("word/document.xml"))

    body = root.find(_W + "body")
    if body is None:
        return []

    # Only top-level paragraphs: the cover tables and the segment-coverage
    # grid are full of text that is not part of the contents.
    paras = [p for p in body.findall(_W + "p")]

    rows = []
    for p in paras:
        pPr = p.find(_W + "pPr")
        style_id = _val(pPr, "pStyle") if pPr is not None else None
        style_id = style_id or "Normal"
        num_id, ilvl = numbering.resolve(p, style_id)
        rows.append({"p": p, "style": style_id, "num": num_id,
                     "ilvl": ilvl or 0, "text": _paragraph_text(p)})

    def is_toc_style(s):
        return s.startswith("ListNumber")

    numbered = [i for i, r in enumerate(rows)
                if r["num"] is not None and is_toc_style(r["style"])]
    if not numbered:
        return []
    first, last = numbered[0], numbered[-1]

    entries: list[TocEntry] = []
    for r in rows[first:last + 1]:
        if not r["text"]:
            continue
        level = min(r["ilvl"], 8)
        if r["num"] is not None:
            marker, font, bullet, hang = numbering.advance(r["num"], r["ilvl"])
            st = _level_style(level, bullet)
            entries.append(TocEntry(text=r["text"], level=level,
                                    marker=marker, marker_font=font,
                                    hanging=hang, bullet=bullet, **st))
        else:
            # An unnumbered line: a banner or a note.  It keeps Word's own
            # look and sits flush, which is where the hand-made deck put it.
            st = dict(size=10.0, bold=False, italic=False, color="000000")
            st.update(_run_format(r["p"]))
            entries.append(TocEntry(text=r["text"], level=level,
                                    marker="", marL=9525, **st))
    return entries


# ── List of Figures / List of Tables ─────────────────────────────────

# Filled during a build by export_native, in the order the deck draws them.
_CAPTURED: dict[str, list[str]] = {"figure": [], "table": []}
_CAPTURING = {"on": False}


def start_capture():
    """Begin recording figure and table captions for this build."""
    _CAPTURED["figure"].clear()
    _CAPTURED["table"].clear()
    _CAPTURING["on"] = True


def stop_capture():
    _CAPTURING["on"] = False


def capture(kind: str, title: str | None):
    """Record one drawn figure or table.  Called from the renderers, so the
    lists describe what the deck REALLY contains - a page dropped because the
    workbook has no volume data takes its figures out of the list with it."""
    if not _CAPTURING["on"] or kind not in _CAPTURED:
        return
    title = (title or "").strip()
    if title:
        _CAPTURED[kind].append(title)


# A panel title on these pages reads "<name> · <how to read it>" - the
# second half is a hint for someone looking AT the chart, and on a list of
# figures it only makes the line wrap.  The name is what belongs in the list.
_CAPTION_TAIL = re.compile(r"\s*[·•–—|]\s.*$")
_CAPTION_MAX = 95              # characters that fit one line of the column


def _clean_caption(title: str) -> str:
    """The figure's NAME: its title without the reading hint after the dot,
    and short enough to stay on one row of a thirty-row table."""
    text = _CAPTION_TAIL.sub("", title).strip(" .")
    text = re.sub(r"\s+", " ", text)
    if len(text) > _CAPTION_MAX:
        text = text[:_CAPTION_MAX - 1].rstrip(" ,;:-") + "…"
    return text


def _numbered(kind: str, word: str) -> list[str]:
    seen, out = set(), []
    for title in _CAPTURED[kind]:
        text = _clean_caption(title)
        if not text:
            continue
        # The same chart can head a continuation slide; number it once.
        key = text.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(f"{word} {len(out) + 1}: {text}")
    return out


def figure_entries() -> list[str]:
    """"Figure 1: ...", in the order the deck draws its charts."""
    return _numbered("figure", "Figure")


def table_entries() -> list[str]:
    """"Table 1: ...", in the order the deck draws its tables."""
    return _numbered("table", "Table")
