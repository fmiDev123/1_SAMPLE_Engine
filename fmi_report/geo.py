"""THE ACTIVE GEOGRAPHY - the one word every page of a report must agree on.

A report is built for ONE geography: `Global`, or one region.  Most renderers
already take that geography as an argument, because they draw its numbers.
Three kinds of text did not:

* the labels that describe the report rather than its data - section names,
  group names, page names, the Agenda, the section dividers, the speaker
  notes.  They live in `registry.py`, are built once at import time, and were
  written for the global report, so they said "Global" whatever the reader had
  selected;
* the prose helpers (`summary.py`, `writeups.py`, `competitors.py`) that sit
  far from any slicer and were phrased "the global market ...";
* a handful of page titles with the word baked into the string.

This module is the argument those places never received.  It is carried per
THREAD, so a Dash worker serving one reader cannot see another's selection,
and it is set in exactly two places: `export_native._export_pptx_native` for
the whole of a deck, and the page callbacks in `app.py` for one render.

Nothing here reads the workbook - the region is whatever the caller selected,
and `data.valid_geos()` has already vetted it by the time it arrives.
"""
from __future__ import annotations

import re
import threading
from contextlib import contextmanager

GLOBAL = "Global"

_state = threading.local()


def current() -> str:
    """The geography being rendered - `"Global"` when nothing is set."""
    return getattr(_state, "region", None) or GLOBAL


def set_current(region: str | None) -> str:
    """Set the active geography for THIS thread; returns the previous one."""
    prev = current()
    _state.region = region or GLOBAL
    return prev


@contextmanager
def active(region: str | None):
    """Render a page (or a whole deck) as `region`, then restore.

        with geo.active("Europe"):
            build_the_deck()
    """
    prev = set_current(region)
    try:
        yield current()
    finally:
        set_current(prev)


def is_global(region: str | None = None) -> bool:
    return (region or current()) == GLOBAL


def word(region: str | None = None) -> str:
    """The geography as a NAME: `"Global"` / `"Europe"`.  Use where the word
    opens a title - "Europe Pricing Assessment"."""
    return region or current()


def adj(region: str | None = None) -> str:
    """The geography INSIDE a sentence: `"global"` / `"Europe"`.

    The global report says "the global market"; a region report says "the
    Europe market" - a proper noun, so it keeps its capital.
    """
    g = region or current()
    return "global" if g == GLOBAL else g


def child_kind(region: str | None = None) -> str:
    """What the geography's children are called: regions under Global,
    countries under a region.  Mirrors `dynamics.facts()["child_kind"]`."""
    return "region" if is_global(region) else "country"


def child_adj(region: str | None = None) -> str:
    """The children as a MODIFIER - "the largest regional market" reads right,
    "the largest region market" does not."""
    return "regional" if is_global(region) else "country"


# A label written for the global report, e.g. "Global Pricing Assessment".
# Matched on the whole word only, so "Globalisation" (or a company called
# "Global Foods") is left alone.
_GLOBAL_RE = re.compile(r"\bglobal\b", re.IGNORECASE)


def _is_region(g) -> bool:
    try:
        from . import data
        return g in data.REGIONS
    except Exception:
        return False


def label(text, region: str | None = None) -> str:
    """Re-point a STATIC label at the active geography.

    "Global Pricing Assessment" → "Europe Pricing Assessment".  Under Global
    (and for any text without the word) it returns the text unchanged, so it
    is safe to wrap every label site with it.

    Only the labels are rewritten - never a slug.  Slugs are built once from
    the registry names and are what the URLs, the deck's internal hyperlinks
    and `pptlock`'s page rulebook are keyed on; renaming those per region
    would break all three.
    """
    g = region or current()
    if g == GLOBAL or not text:
        return text
    s = str(text)
    # a REGION's outlook is read by country, not by region (client,
    # 2026-10-07): 'By Region' -> 'By Country', 'Regional Outlook' -> 'Country
    # Outlook'.  Only for a region - a country has its own wording.
    if _is_region(g):
        s = re.sub(r"\bBy Region\b", "By Country", s)
        s = re.sub(r"\bRegional Outlook\b", "Country Outlook", s)
    if not _GLOBAL_RE.search(s):
        return s
    if g in s:
        # The text already names the region - a title that built itself from
        # the slicer and then appended a static group name, e.g. "Europe ... Market
        # Global Driving Factors".  Substituting would stutter ("Europe ... Europe
        # Driving Factors"), so the stale word is DROPPED instead.  Only the
        # word and the space after it go, which leaves the deck's wide "   ·   "
        # separators alone.
        return re.sub(r"\bglobal\b\s*", "", s, flags=re.IGNORECASE).strip()
    return _GLOBAL_RE.sub(g, s)
