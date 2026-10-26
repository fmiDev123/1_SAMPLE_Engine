"""Speed-ups that change HOW FAST a deck is built, never WHAT is built.

Every change here returns exactly what the code it stands in for returned; the
decks come out byte-for-byte the same (checked part by part against a deck
built without them - see tests in SESSION_NOTES_2026-10-19_speedup.md).

Set the environment variable ``FMI_NO_SPEEDUP=1`` to switch all of it off.

1. ``next_image_partname``.  python-pptx finds the next free ``imageN`` name
   by walking the whole package (every slide, chart and relationship) each
   time a picture is added.  A locked deck adds a picture per padlock and per
   blurred label - well over a thousand times - so the walk dominated the run.
   The walk is done once, the set of used numbers is kept, and it is thrown
   away whenever a relationship is removed (the only way a number can become
   free again), so the name chosen is always the one python-pptx would pick.
2. ``_find_by_sha1``.  Adding a picture also looks for the same image already
   in the package, by hashing every image part - again once per picture.  The
   hashes are read once and kept (same first-match answer, same invalidation).
3. Looking up a shape's placeholder / non-visual properties, and turning
   "p:sp" into "{uri}sp", were each done from scratch millions of times per
   deck.  Same answers, read directly instead of by XPath / re-parsed names.
"""
from __future__ import annotations

import os

_ON = os.environ.get("FMI_NO_SPEEDUP", "").strip() in ("", "0")
_GEN = [0]                                   # bumps when any relationship goes


def _install_image_partname() -> None:
    from pptx.opc.package import _Relationships
    from pptx.opc.packuri import PackURI
    from pptx.package import Package

    orig_pop = _Relationships.pop
    orig_next = Package.next_image_partname

    def pop(self, *a, **k):                  # a rel removed: numbers may free up
        _GEN[0] += 1
        return orig_pop(self, *a, **k)

    def _used(pkg) -> set:
        return {part.partname.idx for part in pkg.iter_parts()
                if part.partname.startswith("/ppt/media/image")
                and part.partname.idx is not None}

    def next_image_partname(self, ext: str):
        cache = getattr(self, "_fmi_img_idx", None)
        if cache is None or cache[0] != _GEN[0]:
            cache = (_GEN[0], _used(self))
            self._fmi_img_idx = cache
        used = cache[1]
        idx = 1
        while idx in used:                   # the lowest number not yet taken
            idx += 1
        used.add(idx)                        # the caller makes the part at once
        return PackURI("/ppt/media/image%d.%s" % (idx, ext))

    _Relationships.pop = pop
    Package.next_image_partname = next_image_partname
    Package._fmi_orig_next_image_partname = orig_next


def _install_image_lookup() -> None:
    from pptx.package import _ImageParts
    from pptx.parts.image import ImagePart

    orig_new = ImagePart.new.__func__

    def _find_by_sha1(self, sha1):
        pkg = self._package
        c = getattr(pkg, "_fmi_sha_idx", None)
        if c is None or c[0] != _GEN[0]:
            found = {}
            for part in self:                # same walk order as python-pptx
                if not hasattr(part, "sha1"):
                    continue                 # unknown type (SVG): skipped there too
                found.setdefault(part.sha1, part)
            c = (_GEN[0], found)
            pkg._fmi_sha_idx = c
        return c[1].get(sha1)

    def new(cls, package, image):
        part = orig_new(cls, package, image)
        c = getattr(package, "_fmi_sha_idx", None)
        if c is not None and c[0] == _GEN[0]:
            c[1].setdefault(part.sha1, part)
        return part

    _ImageParts._find_by_sha1 = _find_by_sha1
    ImagePart.new = classmethod(new)


def _install_shape_lookups() -> None:
    import types
    from pptx.oxml import ns
    from pptx.oxml.shapes.shared import BaseShapeElement

    _P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    NVPR, PH = _P + "nvPr", _P + "ph"

    def _first_child(el):
        """``./*[1]``: the first child that is an element (not a comment)."""
        for ch in el:
            if isinstance(ch.tag, str):
                return ch
        return None

    def ph(self):
        first = _first_child(self)
        if first is None:
            return None
        nvpr = first.find(NVPR)
        return None if nvpr is None else nvpr.find(PH)

    def _nvXxPr(self):
        first = _first_child(self)
        if first is None:
            raise IndexError("list index out of range")
        return first

    BaseShapeElement.ph = property(ph, doc=BaseShapeElement.ph.__doc__)
    BaseShapeElement._nvXxPr = property(_nvXxPr)

    # qn("p:sp"): the same string every time, so build it once
    orig = types.FunctionType(ns.qn.__code__, ns.qn.__globals__, "qn")
    cache = {}

    def qn(tag, _cache=cache, _orig=orig):
        try:
            return _cache[tag]
        except KeyError:
            _cache[tag] = r = _orig(tag)
            return r

    ns.qn.__code__ = qn.__code__
    ns.qn.__defaults__ = qn.__defaults__


def install() -> None:
    if not _ON or getattr(install, "done", False):
        return
    install.done = True
    try:
        _install_image_partname()
        _install_image_lookup()
        _install_shape_lookups()
    except Exception:                        # never stop a build over a speed-up
        pass


install()
