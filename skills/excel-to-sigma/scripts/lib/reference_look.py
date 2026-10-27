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
HEX = re.compile(r"#[0-9a-fA-F]{3,8}")


def workbook_id(text):
    """The id ``GET /v2/workbooks/{id}/spec`` accepts, from a bare id / UUID or a workbook URL.

    From a URL, the path segment right after ``/workbook/`` is used (nothing after it: ``/page/..``, ``?..``,
    ``#..``). A UUID inside it wins; otherwise the text after its last ``-`` (a ``Golden-Orders-<id>`` slug is
    rejected by the API, its trailing id is accepted).
    """
    text = (text or "").strip()
    if not text:
        return None
    text = re.split(r"[?#]", text, 1)[0]
    found = re.search(r"/workbook/([^/]+)", text)
    segment = found.group(1) if found else text.rstrip("/").rsplit("/", 1)[-1]
    uuid = UUID.search(segment)
    if uuid:
        return uuid.group(0).lower()
    return segment.rsplit("-", 1)[-1] or None


def _dict(value):
    return value if isinstance(value, dict) else {}


def _document(spec):
    spec = _dict(spec)
    for key in ("document", "contents"):
        if isinstance(spec.get(key), dict):
            return spec[key]
    return spec


def _kind_of(element):
    kind = element.get("kind")
    kind = kind if isinstance(kind, str) else ""
    if kind in ("kpi-chart", "container", "table", "pivot-table"):
        return next(group for group, kinds in CARD_KINDS.items() if kind in kinds)
    return "chart" if kind.endswith("-chart") else None


def _hidden_element_ids(document):
    """Ids of elements laid out on hidden pages (``visibility: hidden``), through ``layout``'s ``<Page id>``."""
    hidden_pages = set()
    for page in document.get("pages") if isinstance(document.get("pages"), list) else []:
        visibility = _dict(page).get("visibility")
        if visibility == "hidden" or (isinstance(visibility, dict) and "hidden" in json.dumps(visibility)):
            hidden_pages.add(page.get("id"))
    layout = document.get("layout")
    if not hidden_pages or not isinstance(layout, str):
        return set()
    hidden = set()
    for match in re.finditer(r'<Page\b[^>]*\bid="([^"]*)"[^>]*>(.*?)</Page>', layout, re.S):
        if match.group(1) in hidden_pages:
            hidden.update(re.findall(r'elementId="([^"]*)"', match.group(2)))
    return hidden


def _most_common(items, minimum_share=0.5):
    """The most common non-empty dict of ``items`` (empty / non-dict entries count against it), or ``None``
    when it does not cover at least ``minimum_share`` of them."""
    items = list(items)
    if not items:
        return None
    counts = Counter(json.dumps(i, sort_keys=True) for i in items if isinstance(i, dict) and i)
    if not counts:
        return None
    value, count = counts.most_common(1)[0]
    return json.loads(value) if count / len(items) >= minimum_share else None


def _choose(visible, everything):
    """Most common over the visible elements; when none of them is styled, over the styled ones anywhere."""
    found = _most_common(visible)
    if found is not None or any(isinstance(v, dict) and v for v in visible):
        return found
    return _most_common([v for v in everything if isinstance(v, dict) and v])


def is_hex(value):
    return isinstance(value, str) and bool(HEX.fullmatch(value))


def _number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def safe_fields(source, fields):
    """The listed styling fields of ``source`` that are safe to write into a style (a hex colour, a number, a
    known weight); anything else is dropped."""
    out = {}
    for field in fields:
        value = _dict(source).get(field)
        if field == "color" and isinstance(value, str) and HEX.fullmatch(value):
            out[field] = value
        elif field == "fontSize" and _number(value):
            out[field] = value
        elif field == "fontWeight" and value in ("bold", "normal"):
            out[field] = value
    return out


def _span(body):
    """``{size, color}`` of the first ``<span style>`` of a text body (px sizes only)."""
    found = re.search(r'<span style="([^"]*)"', body) if isinstance(body, str) else None
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
    """The reference look of a workbook spec; see the module docstring for the shape.

    Only elements on visible pages vote (a hidden data page full of bare tables must not outvote the styled
    page people see); if nothing visible is styled, the styled elements anywhere decide.
    """
    document = _document(spec)
    everything = [e for e in document.get("elements") or [] if isinstance(e, dict)] \
        if isinstance(document.get("elements"), list) else []
    hidden = _hidden_element_ids(document)
    shown = [e for e in everything if e.get("id") not in hidden]
    settings = _dict(document.get("settings"))
    theme = settings.get("theme")
    theme = copy.deepcopy(theme) if isinstance(theme, dict) and theme else None

    styles = {}
    for group in ("container", "kpi-chart", "chart", "table"):
        styles[group] = _choose([e.get("style") for e in shown if _kind_of(e) == group],
                                [e.get("style") for e in everything if _kind_of(e) == group])

    def kpis(elements):
        return [e for e in elements if e.get("kind") == "kpi-chart"]
    kpi = {"name": _choose([safe_fields(e.get("name"), KPI_NAME_FIELDS) or None for e in kpis(shown)],
                           [safe_fields(e.get("name"), KPI_NAME_FIELDS) or None for e in kpis(everything)]),
           "value": _choose([safe_fields(e.get("value"), KPI_VALUE_FIELDS) or None for e in kpis(shown)],
                            [safe_fields(e.get("value"), KPI_VALUE_FIELDS) or None for e in kpis(everything)])}

    colors = _dict(_dict(_dict(theme).get("overrides")).get("colors"))
    texts = [e for e in shown if e.get("kind") == "text"] or [e for e in everything if e.get("kind") == "text"]
    title = subline = None
    if texts:
        first = texts[0].get("body")
        first = first if isinstance(first, str) else ""
        span = _span(first)
        heading = re.match(r"\s*(#{1,6})\s", first)
        if span.get("size") or heading:
            title = {**span}
            if heading and not span.get("size"):
                title["heading"] = len(heading.group(1))
        if len(texts) > 1:
            sub = _span(texts[1].get("body"))
            if sub.get("size"):
                subline = sub

    def presets(elements):
        return [_dict(e.get("tableStyle")).get("preset") for e in elements if e.get("kind") in ("table", "pivot-table")]
    counted = Counter(p for p in presets(shown) if isinstance(p, str) and p) or \
        Counter(p for p in presets(everything) if isinstance(p, str) and p)
    preset = counted.most_common(1)
    text_color = colors.get("text")
    return {"version": LOOK_VERSION, "theme": theme, "styles": styles, "kpi": kpi,
            "text": {"title": title, "subline": subline,
                     "color": text_color if isinstance(text_color, str) and HEX.fullmatch(text_color) else None},
            "tablePreset": preset[0][0] if preset else None}
