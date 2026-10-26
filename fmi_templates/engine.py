"""
Render engine and QC for FMI slide templates.

    from fmi_templates.engine import render_slide
    out = render_slide("G03", ctx, mask=True)      # dict: element -> text or list of texts

Masking: every {v_*} placeholder renders as "XX" when mask=True (sample decks).
Missing placeholders raise MissingKeys so a slide never ships with a gap.
"""
import re
import string

from .lexicon import lexicon_for
from .rules import derive, item_rules
from .slide_templates import SLIDES, COMMON

SLIDE_INDEX = {s["id"]: s for s in SLIDES}
MASK = "XX"
_fmt = string.Formatter()


class MissingKeys(KeyError):
    pass


def _fields(text):
    return [f for _, f, _, _ in _fmt.parse(text) if f]


def render_text(text, ctx, mask=True, _depth=0):
    """Fill placeholders; rule phrases may contain placeholders, so resolve up to 4 levels."""
    if not isinstance(text, str):
        return text
    missing = []
    out = []
    for literal, field, spec, conv in _fmt.parse(text):
        out.append(literal)
        if field is None:
            continue
        if field.startswith("v_") and mask:
            out.append("X.X" if re.search(r"(mult|ratio|vs_second|_disp)$", field) else MASK)
            continue
        if field not in ctx:
            missing.append(field)
            out.append("{" + field + "}")
            continue
        val = ctx[field]
        if isinstance(val, (list, tuple)):
            val = ", ".join(map(str, val))
        out.append(format(val, spec) if spec else str(val))
    res = "".join(out)
    if missing:
        raise MissingKeys(sorted(set(missing)))
    if "{" in res and _depth < 4 and _fields(res):
        return render_text(res, ctx, mask, _depth + 1)
    return clean(res)


def clean(text):
    text = text.replace("–", "-").replace("—", "-")      # house rule: no en or em dashes
    text = re.sub(r"[ \t]{2,}", lambda m: m.group(0) if len(m.group(0)) > 3 else " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"\.\.(?!\.)", ".", text)
    return text.strip()


# Country names that take "the" in running text. Extend as needed.
THE_NAMES = {"U.S.", "U.K.", "UAE", "Netherlands", "Philippines", "Czech Republic", "Dominican Republic",
             "United States", "United Kingdom", "United Arab Emirates", "Rest of the World"}
NAME_KEYS = ("geo", "cty_", "child_", "focus_country", "parent_geo", "c_geo_")


def _articles(ctx):
    out = {}
    for k, v in ctx.items():
        if isinstance(v, str) and (k == "geo" or k.startswith(NAME_KEYS)) and not k.startswith("child_label"):
            needs = v in THE_NAMES
            out[k + "_the"] = ("the " + v) if needs else v
            out[k + "_The"] = ("The " + v) if needs else v
    return out


def build_ctx(ctx):
    """Merge lexicon, article forms and derived phrases into the pipeline context."""
    full = dict(lexicon_for(ctx["industry"]))
    full.update(ctx)
    if "child_label" in full and "child_label_adj" not in full:
        full["child_label_adj"] = {"region": "regional", "country": "country"}.get(full["child_label"], full["child_label"])
    if "child_label_adj" in full:
        full["child_label_adj_cap"] = full["child_label_adj"][:1].upper() + full["child_label_adj"][1:]
    if "vc_top_stage" in full:
        full["vc_top_stage_lc"] = full["vc_top_stage"].lower()
    full.update(_articles(full))
    full.update(derive(full))
    return full


def _render_value(val, ctx, mask):
    if isinstance(val, dict) and "repeat" in val:
        items = ctx.get(val["repeat"])
        if items is None:
            raise MissingKeys([val["repeat"]])
        res = []
        for i, item in enumerate(items, 1):
            item = dict(item) if isinstance(item, dict) else {"item": item}
            item.setdefault("rank", i)
            local = dict(ctx)
            local.update(item)
            local.update(item_rules(val["repeat"], item, ctx))
            res.append(render_text(val["text"], local, mask))
        return res
    if isinstance(val, dict):
        return {render_text(k, ctx, mask): _render_value(v, ctx, mask) for k, v in val.items()}
    if isinstance(val, (list, tuple)):
        res = []
        for v in val:
            r = _render_value(v, ctx, mask)
            res.extend(r if isinstance(v, dict) and "repeat" in v else [r])
        return res
    return render_text(val, ctx, mask)


SKIP_ELEMENTS = {"rule_note", "likelihood_rule", "private_variant", "no_moves_variant", "empty_row",
                 "subtitle_value_only", "row_real_dev", "toc_patterns", "figure_patterns", "table_patterns"}


def render_slide(sid, ctx, mask=True, include_notes=False):
    """Render one slide type. ctx must already be passed through build_ctx()."""
    s = SLIDE_INDEX[sid]
    out = {"id": sid, "name": s["name"]}
    errors = {}
    for key in ("title", "subtitle"):
        try:
            out[key] = render_text(s[key], ctx, mask)
        except MissingKeys as e:
            errors[key] = e.args[0]
    for key, val in list(s["elements"].items()) + list(COMMON.items()):
        if key in SKIP_ELEMENTS and not include_notes:
            continue
        try:
            out[key] = _render_value(val, ctx, mask)
        except MissingKeys as e:
            errors[key] = e.args[0]
    if errors:
        out["_missing"] = errors
    return out


# ------------------------------------------------------------------------------------------ QC
_DUP = re.compile(r"\b(\w+)\s+\1\b", re.I)
_NUM_LEAK = re.compile(r"(?<![A-Za-z])[-+]?\d+(\.\d+)?\s?(%|BPS|x\b)")


def qc_text(text, mask=True):
    issues = []
    if "{" in text or "}" in text:
        issues.append("unresolved placeholder")
    if "–" in text or "—" in text:
        issues.append("en or em dash")
    m = _DUP.search(text)
    if m and m.group(1).lower() not in {"that", "had"}:
        issues.append(f"duplicated word '{m.group(1)}'")
    leaks = [m.group(0) for m in _NUM_LEAK.finditer(text) if m.group(0) not in ("100%",)]
    if mask and leaks:
        issues.append(f"figure leaked in masked deck: '{leaks[0]}'")
    if re.search(r"\(\s*\)", text):
        issues.append("empty brackets")
    if re.search(r"\b(hospital|clinic|clinical)\b", text, re.I) and "_food_" in text:
        issues.append("industry wording mismatch")
    return issues


def qc_slide(rendered, mask=True, banned_terms=()):
    issues = []

    def walk(k, v):
        if isinstance(v, str):
            for i in qc_text(v, mask):
                issues.append((k, i, v[:90]))
            for term in banned_terms:
                if re.search(rf"\b{re.escape(term)}\b", v, re.I):
                    issues.append((k, f"banned term '{term}'", v[:90]))
        elif isinstance(v, list):
            for x in v:
                walk(k, x)
        elif isinstance(v, dict):
            for kk, x in v.items():
                walk(f"{k}.{kk}", x)

    for k, v in rendered.items():
        if k in ("id", "name", "_missing"):
            continue
        walk(k, v)
    if "_missing" in rendered:
        issues.append(("_missing", "missing placeholders", str(rendered["_missing"])))
    return issues


BANNED_BY_INDUSTRY = {
    "food": ["hospital", "clinic", "clinical", "patient", "formulary", "aftermarket"],
    "consumer": ["hospital", "clinical", "patient", "formulary"],
    "healthcare": ["shelf", "flavour", "grocery"],
    "pharma": ["shelf", "flavour", "grocery", "aftermarket"],
    "industrial": ["patient", "clinical", "flavour", "grocery"],
}
