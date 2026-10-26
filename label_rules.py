"""label_rules.py - shared label rules of the report core (one place, every
format and every domain).

Parenthetical rule (client, 2026-10-14): a name keeps a "(...)" only when the
brackets hold an ABBREVIATION that is not longer than the name itself -
'Electronics For Imaging, Inc. (EFI)' stays, 'Slotted Shipping Boxes (RSC,
HSC, FOL and other slotted shipping containers)' becomes 'Slotted Shipping
Boxes'.  Anything else in brackets is trimmed.
"""
from __future__ import annotations

import re

_PAREN = re.compile(r"\s*\(([^()]*)\)")


def _is_abbrev_token(tok: str) -> bool:
    """'EFI', 'RSC', 'IoT', '3D', 'R&D', 'U.S.' - short, no spaces, at least
    two capitals (or a digit with a capital); a plain word is not."""
    t = tok.strip()
    if not t or " " in t or len(t) > 8:
        return False
    caps = sum(ch.isupper() for ch in t)
    return caps >= 2 or (caps >= 1 and any(ch.isdigit() for ch in t))


def is_abbreviation(text: str) -> bool:
    """One abbreviation, or a comma/slash list of abbreviations ('RSC, HSC')."""
    toks = [x for x in re.split(r"\s*[,;/]\s*", str(text).strip()) if x]
    return bool(toks) and all(_is_abbrev_token(x) for x in toks)


def trim_parenthetical(label) -> str:
    """Drop every '(...)' that is not an abbreviation, or that is longer than
    the rest of the name.  A name that is nothing but a bracket is kept."""
    s = "" if label is None else str(label)
    if "(" not in s:
        return s
    base = _PAREN.sub("", s).strip(" ,;-")
    if not base:
        return s

    def keep(m):
        inner = m.group(1).strip()
        return m.group(0) if is_abbreviation(inner) and len(inner) <= len(base) \
            else ""
    out = _PAREN.sub(keep, s)
    return re.sub(r"\s{2,}", " ", out).strip(" ,;-") or s


def trim_labels(labels):
    """{original: trimmed} for a group of sibling labels.  A trim that would
    make two different siblings read the same is not applied to them - that
    would merge two sub-segments into one."""
    uniq = list(dict.fromkeys(str(x) for x in labels))
    new = {x: trim_parenthetical(x) for x in uniq}
    seen: dict[str, list[str]] = {}
    for x, n in new.items():
        seen.setdefault(n.casefold(), []).append(x)
    for n, xs in seen.items():
        if len(xs) > 1:
            for x in xs:
                new[x] = x
    return new
