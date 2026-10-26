"""THE TEMPLATE SWITCH - one place that knows two templates exist.

The report ships two visual templates, Fact.MR and FMI, and either can be
selected at runtime from the dashboard or passed to the PowerPoint export.

    theming.apply("fmi")        # make FMI the active template
    theming.current()           # -> "fmi"
    with theming.use("factmr"): # temporarily, then restore
        build_the_deck()

WHY A REBINDING ENGINE AND NOT A GLOBAL LOOKUP
    Every renderer reads its colours as plain module globals, imported once at
    start-up:

        from .config import ACCENT, HEADER_BG, TEXT      # components.py
        _C_HEADER = _rgb(HEADER_BG)                      # export_native.py

    That is 7 modules holding their own copy of the VALUE, plus ~65 constants
    derived from those values at import time.  Mutating `config` therefore
    changes nothing.  Rewriting all several-hundred usage sites to go through a
    live object would touch every chart in the codebase and risk the one thing
    that must not regress: the Fact.MR report.

    So the switch works the other way round.  `apply()` walks a table of
    consumer modules, rebinds the token names in each module's globals, then
    calls that module's `_retheme()` hook to re-derive its own constants.  The
    renderers keep reading plain globals and never learn that templates exist -
    which is exactly the "no `if fmi:` scattered through the codebase" rule.

CONCURRENCY
    Which template a caller WANTS is per thread (`current()`), the way
    `geo.py` carries the active geography.  Which template is actually BOUND
    into the consumer modules is process-wide, because the rebinding is.

    So a render must hold `LOCK` across the bind and the drawing both, or
    another thread rebinds underneath it:

        with theming.rendering(tpl):        # page renders
            body = draw_the_page()

        with theming.use(tpl):              # deck builds - also restores
            build_the_deck()

    Bare `apply()` selects a template and returns, releasing LOCK; it is for
    start-up and for the wordmark sync, NOT for anything that then renders.

    This serialises rendering within one process rather than giving each
    viewer their own palette - the renderers read plain module globals, and
    making those per-thread would mean rewriting every chart in the codebase.
    Capacity therefore comes from running more processes, not more threads.
"""
from __future__ import annotations

import importlib
import threading
from contextlib import contextmanager
from pathlib import Path

from . import config, theme_factmr, theme_fmi

# ── The templates ────────────────────────────────────────────────────
# A template IS its theme module: both expose the same token names, so there is
# nothing to configure per template beyond pointing at the module.
TEMPLATES: dict[str, object] = {
    theme_factmr.NAME: theme_factmr,
    theme_fmi.NAME: theme_fmi,
}
DEFAULT = theme_factmr.NAME          # Fact.MR stays the default, as shipped
ORDER = [theme_factmr.NAME, theme_fmi.NAME]

LABELS = {name: mod.LABEL for name, mod in TEMPLATES.items()}

_ASSETS = Path(__file__).resolve().parent / "assets"

# Serialises "apply theme -> render" so a concurrent request cannot swap the
# palette mid-page.  Re-entrant: the export path takes it around a whole deck
# and the page renderer may take it again inside.
LOCK = threading.RLock()

# ── Which modules hold their own copy of the tokens ──────────────────
# Derived from the actual `from .config import ...` lines.  A module is listed
# once; `apply()` only rebinds names the module already defines, so adding a
# token here that a module never imported is a no-op rather than an injection.
CONSUMERS = (
    "config",
    "layout",
    "components",
    "figures",
    "export_native",
    "export",
    "app",
    "templates_main",
    "templates_country",
    "templates_competitor",
    "templates_dynamics",
    "templates_value_chain",
    "competitors",
    "dynamics",
    "value_chain",
    "summary",
    "writeups",
    "pricing_drivers",
    "metrics",
)


def tokens(mod=None) -> tuple[str, ...]:
    """Every design token a theme module publishes.

    Everything public and UPPER_CASE, plus the component-metric helpers.  Both
    themes define the same set, so this is the contract the renderers see.
    """
    mod = mod or TEMPLATES[DEFAULT]
    names = [n for n in vars(mod)
             if not n.startswith("_") and (n.isupper() or n in _HELPERS)]
    return tuple(sorted(names))


_HELPERS = ("line_px", "wrap_lines", "stretch_factor", "table_row_h",
            "table_head_h", "kpi_value_size", "kpi_height", "kpi_value_px",
            "panel_title_height", "summary_height", "NAME", "LABEL")

# Tokens that are NOT the theme's business: `config` owns the dataset, the
# report years, the unit labels and the geo map, and a theme must never clobber
# them.  (No theme defines these today; the guard keeps it that way.)
PROTECTED = frozenset({
    "DATA_FILE", "YEARS", "HIST_START", "HIST_END", "EST_YEAR", "MID_YEAR",
    "END_YEAR", "HIST_YEARS", "FORE_YEARS", "UNIT_LABELS", "GEO_ISO3",
    "TEMPLATE",
})

# WHICH TEMPLATE THIS THREAD ASKED FOR.  Carried per thread for the reason
# `geo.py` carries the active geography per thread: a worker serving one
# reader must not see another reader's selection.
_state = threading.local()

# WHICH TEMPLATE'S TOKENS ARE ACTUALLY BOUND into the consumer modules right
# now.  Process-wide, because the rebinding is; only ever touched under LOCK.
#
# The two can disagree - thread A selects FMI, thread B rebinds to Fact.MR -
# and closing that gap is what `rendering()` below is for: it holds LOCK
# across the bind AND the render, so nothing can rebind underneath a
# half-drawn page.  Selecting a template outside that context records this
# thread's preference; it does not promise the globals still agree by the
# time you read them.
_bound: str | None = None


def _startup() -> str:
    """The template `config` chose at import time from REPORT_TEMPLATE.

    Deliberately does not go through `normalise()`: that falls back to
    `current()`, which would call straight back into here.
    """
    key = getattr(config, "TEMPLATE", DEFAULT)
    return key if key in TEMPLATES else DEFAULT


def current() -> str:
    """The template name ('factmr' / 'fmi') THIS THREAD has selected."""
    return getattr(_state, "active", None) or _startup()


def label(name: str | None = None) -> str:
    return LABELS.get(name or current(), LABELS[DEFAULT])


def theme(name: str | None = None):
    """The theme module for `name` (the active one by default)."""
    return TEMPLATES[normalise(name or current())]


def normalise(name: str | None) -> str:
    """Accept 'FMI', 'fmi', 'Fact.MR', 'factmr', None -> a known key."""
    if not name:
        return current()
    key = str(name).strip().lower().replace(".", "").replace(" ", "")
    if key in TEMPLATES:
        return key
    for cand, mod in TEMPLATES.items():
        if key == mod.LABEL.lower().replace(".", "").replace(" ", ""):
            return cand
    return DEFAULT


def logo_path(name: str | None = None) -> Path:
    """Absolute path to the active template's wordmark, with its fallback."""
    t = theme(name)
    p = _ASSETS / t.LOGO_FILE
    if not p.exists():
        p = _ASSETS / getattr(t, "LOGO_FALLBACK", t.LOGO_FILE)
    return p


def asset_url(name: str | None = None) -> str:
    """Dash asset URL for the active template's wordmark."""
    return "/assets/" + logo_path(name).name


def _bind(key: str) -> None:
    """Rebind every token in every consumer module to template `key`.

    THE CALLER MUST HOLD `LOCK`.  A no-op when `key` is already bound, which
    is what makes it cheap enough to call on every page render.
    """
    global _bound
    if _bound == key:
        return
    src = TEMPLATES[key]
    # Publish the new template BEFORE rebinding.  The `_retheme()` hooks call
    # back into this module (theming.logo_path(), theming.theme()) to resolve
    # the wordmark, its aspect ratio and the copyright line; if the outgoing
    # template were still the published one those hooks would quietly pick up
    # the OLD brand assets, which is exactly how the FMI deck once ended up
    # embedding the Fact.MR wordmark.
    config.TEMPLATE = key
    _bound = key
    values = {n: getattr(src, n) for n in tokens(src) if n not in PROTECTED}
    for modname in CONSUMERS:
        try:
            mod = importlib.import_module(f".{modname}", __package__)
        except Exception:
            continue                       # module not present / not built
        g = vars(mod)
        for n, v in values.items():
            # `config` re-exports the whole theme, so it takes every token;
            # every other module only takes what it already imported.
            if modname == "config" or n in g:
                g[n] = v
            # A module that already had a name of its own for a token imports
            # it aliased, e.g. `PP_COLORS as _T_PP_COLORS`.  The alias is
            # private, so it is not a token and would otherwise keep its
            # import-time value forever - which is exactly how the FMI deck
            # kept a few Fact.MR colours.  Rebind it as well.
            alias = "_T_" + n
            if alias in g:
                g[alias] = v
        hook = g.get("_retheme")
        if callable(hook):
            hook()


def apply(name: str | None = None) -> str:
    """Select `name` for this thread and bind its tokens.

    Kept for callers that are not rendering (start-up, the wordmark sync).
    A RENDER should use `rendering()` instead: `apply` releases LOCK on the
    way out, so by the time the caller draws anything another thread may
    already have rebound the tokens underneath it.
    """
    key = normalise(name)
    with LOCK:
        _state.active = key
        _bind(key)
    return key


@contextmanager
def rendering(name: str | None = None):
    """Hold LOCK with `name`'s tokens bound - the one safe way to render.

        with theming.rendering(tpl):
            body = draw_the_page()

    Two viewers on different templates therefore take turns rather than
    corrupting each other's palette.  This serialises rendering inside one
    process; capacity comes from running more of them.

    Unlike `use()` this does NOT restore the previous template on the way
    out.  Nothing reads the tokens outside a render context, and the next
    render binds whatever it needs, so restoring would only buy a second
    rebind per page.
    """
    with LOCK:
        key = normalise(name)
        _state.active = key
        _bind(key)
        yield key


@contextmanager
def use(name: str | None):
    """Temporarily switch template, holding LOCK, then restore.

    Used by the PowerPoint export so a deck can be built in one template while
    the dashboard sits on the other.
    """
    with LOCK:
        before = current()
        target = normalise(name)
        _state.active = target
        _bind(target)
        try:
            yield target
        finally:
            _state.active = before
            _bind(before)
