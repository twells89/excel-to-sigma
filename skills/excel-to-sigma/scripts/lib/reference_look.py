"""Extract the *look* of an existing Sigma workbook ("make it look like Golden Orders").

``extract_look(spec_document)`` is a pure function over a workbook spec (a ``GET /v2/workbooks/{id}/spec``
response, or its ``document`` / ``contents`` part). It keeps the look only - theme, card styles, KPI type, text
type and the table preset - never the arrangement or any data. The table-report builder applies it when the plan
carries ``presentation.look``; the Excel file's own decisions (number formats, header fill, row order, titles)
still win.

Look (JSON)::

    {"version": 1,
     "theme":  <settings.theme copied exactly> | null,
     "styles": {"container": {...}|null, "kpi-chart": ..., "chart": ..., "table": ...},   # most common per kind
     "kpi":    {"name": {color, fontSize, fontWeight}|null, "value": {fontSize, fontWeight, color}|null},
     "text":   {"title": {size, color, heading}|null, "subline": {size, color}|null, "color": <theme text>|null},
     "tablePreset": "presentation"|null}
"""
from __future__ import annotations

import copy
import json
import re
from collections import Counter

LOOK_VERSION = 1
CARD_KINDS = {"container": ("container",), "kpi-chart": ("kpi-chart",), "table": ("table", "pivot-table")}
KPI_NAME_FIELDS = ("color", "fontSize", "fontWeight")
KPI_VALUE_FIELDS = ("fontSize", "fontWeight", "color")
UUID = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")


def workbook_id(text):
    """A workbook id from a bare id or a Sigma workbook URL."""
    found = UUID.findall(text or "")
    if found:
        return found[-1].lower()
    tail = (text or "").rstrip("/").rsplit("/", 1)[-1].split("?")[0]
    return tail or None


def _document(spec):
    for key in ("document", "contents"):
        if isinstance(spec.get(key), dict):
            return spec[key]
    return spec


def _kind_of(element):
    kind = element.get("kind", "")
    if kind in ("kpi-chart", "container", "table", "pivot-table"):
        return next(group for group, kinds in CARD_KINDS.items() if kind in kinds)
    return "chart" if kind.endswith("-chart") else None


def _most_common(items, minimum_share=0.5):
    """The most common non-empty dict of ``items`` (None entries count against it), or ``None`` when it does
    not cover at least ``minimum_share`` of them."""
    if not items:
        return None
    counts = Counter(json.dumps(i, sort_keys=True) for i in items if i)
    if not counts:
        return None
    value, count = counts.most_common(1)[0]
    return json.loads(value) if count / len(items) >= minimum_share else None


def _pick(source, fields):
    return {f: source[f] for f in fields if isinstance(source, dict) and source.get(f) is not None}


def _span(body):
    """``{size, color}`` of the first ``<span style>`` of a text body (px sizes only)."""
    found = re.search(r'<span style="([^"]*)"', body or "")
    out = {}
    if found:
        size = re.search(r"font-size:\s*(\d+(?:\.\d+)?)px", found.group(1))
        color = re.search(r"(?<![-\w])color:\s*(#[0-9A-Fa-f]{3,8})", found.group(1))
        if size:
            out["size"] = float(size.group(1)) if "." in size.group(1) else int(size.group(1))
        if color:
            out["color"] = color.group(1)
    return out


def extract_look(spec):
    """The reference look of a workbook spec; see the module docstring for the shape."""
    document = _document(spec or {})
    elements = [e for e in document.get("elements") or [] if isinstance(e, dict)]
    theme = (document.get("settings") or {}).get("theme")
    theme = copy.deepcopy(theme) if isinstance(theme, dict) and theme else None

    styles = {}
    for group in ("container", "kpi-chart", "chart", "table"):
        members = [e for e in elements if _kind_of(e) == group]
        styles[group] = _most_common([e.get("style") for e in members])

    kpis = [e for e in elements if e.get("kind") == "kpi-chart"]
    names = [_pick(e["name"], KPI_NAME_FIELDS) if isinstance(e.get("name"), dict) else None for e in kpis]
    values = [_pick(e.get("value"), KPI_VALUE_FIELDS) or None for e in kpis]
    kpi = {"name": _most_common(names), "value": _most_common(values)}

    colors = ((theme or {}).get("overrides") or {}).get("colors") or {}
    texts = [e for e in elements if e.get("kind") == "text"]
    title = subline = None
    if texts:
        first = texts[0]["body"] or ""
        span = _span(first)
        heading = re.match(r"\s*(#{1,6})\s", first)
        if span.get("size") or heading:
            title = {**span}
            if heading and not span.get("size"):
                title["heading"] = len(heading.group(1))
        if len(texts) > 1:
            sub = _span(texts[1]["body"])
            if sub.get("size"):
                subline = sub
    presets = [(e.get("tableStyle") or {}).get("preset") for e in elements if e.get("kind") in ("table", "pivot-table")]
    preset = Counter(p for p in presets if p).most_common(1)
    return {"version": LOOK_VERSION, "theme": theme, "styles": styles, "kpi": kpi,
            "text": {"title": title, "subline": subline, "color": colors.get("text")},
            "tablePreset": preset[0][0] if preset else None}


def span_style(size=None, color=None):
    """A ``style`` attribute value with only properties Sigma accepts, or ``""``."""
    parts = []
    if size:
        parts.append(f"font-size: {size}px")
    if color:
        parts.append(f"color: {color}")
    return "; ".join(parts)
