"""Presentation styling for the table-report Report page.

Ported (read-only, no code shared) from the sigma-skills ``sigma-workbooks`` skill at commit 95e67b7:

* ``reference/specification/styling.md`` - the vetted palette, Recipes 1-6 (hero header strip, KPI card row,
  section header, divider, categorical chart colors, number formatting), the anti-pattern checklist and the
  *Composition principles*;
* ``scripts/lib/styling.rb`` - semantic roles -> verified hexes (``DEFAULT_ROLES``), ``header``,
  ``section_card``, ``kpi_accent``, ``chart_color``, ``format_for`` and the ``SURFACES`` map.

styling.md's own rule applies: these recipes never override a migration's source fidelity. Source number
formats, row order and header fills win; the recipes only fill what Excel does not define (page chrome, chart
colors, KPI cards, spacing).

Every field emitted below is one the offline Sigma public OpenAPI defines (``components.schemas.Table`` /
``CommonElement``); ``SURFACES`` names each shape with its OpenAPI location so a surface that turns out to be
rejected live can be switched off in one place and the builder simply omits it. Role names are authoring-only:
only resolved hex / ``{kind: theme, ref}`` values are ever written into a Sigma color field.
"""
from __future__ import annotations

import colorsys
import math

# ---------------------------------------------------------------- surfaces (allow-list of emitted shapes)
SURFACES = {
    # name: OpenAPI location of the shape (and whether sigma-workbooks styling.md live-verified it)
    "container_style": True,       # container.style {backgroundColor, borderRadius} + a text child  (verified)
    "text_body_spans": True,       # text.body <span style=color> / <p class=p-small>, verticalAlign   (verified)
    "kpi_name_object": True,       # kpi-chart.name {text, color}; value {columnId, fontSize}      (verified)
    "kpi_card_style": True,        # kpi-chart.style {backgroundColor, borderColor, borderWidth, borderRadius}
    "chart_card_style": True,      # bar/line/donut .style (same card)
    "chart_color_single": True,    # bar/line .color {by: single, value}                           (verified)
    "categorical_scheme": True,    # document.settings.theme.overrides.categoricalScheme           (verified)
    "format_string": True,         # columns[].format {kind, formatString (d3), prefix, suffix}    (verified)
    "table_text_styles": True,     # table/pivot .tableStyle.textStyles.{header, columnHeader}
    "conditional_single": True,    # table/pivot .conditionalFormats[{type: single, condition: <, value, style}]
    "divider": True,               # divider element                                               (verified)
    "order_group_sort": True,      # hidden `Min(<order If>)` calculation first in calculations + groupings[].sort
                                   # (grouping sort is accepted live though the offline OpenAPI omits it)
    "chart_hidden_column": True,   # chart columns[].hidden (sort-only order column)
    "chart_label_angle": True,     # bar-chart xAxis.format.labels.labelAngle: 0 (12 categories or fewer)
    "chart_list_filter": True,     # chart filters[{kind: list, mode: include}] on the category column
    "kpi_text_card": True,         # container card + text with {{Sum([A/x]) / Sum([B/y]) | fmt}} for a KPI that
                                   # spans several source elements (a kpi-chart has ONE source)
}

# ---------------------------------------------------------------- palette (styling.rb DEFAULT_ROLES)
ROLES = {
    "canvas": "#FFFFFF", "ink": "#0F172A", "body": "#1E293B", "muted": "#64748B", "hairline": "#E2E8F0",
    "fill": "#FFFFFF", "series": "#2563EB", "series-mid": "#0EA5E9", "series-grey": "#64748B",
    "primary": "#2563EB", "tint": "#1E3A8A", "edge": "#FFFFFF",
}
# categorical slots 2..6 (styling.md "Additional verified categorical hues")
CATEGORICAL_MID = ("#14B8A6", "#F59E0B", "#8B5CF6", "#EF4444", "#10B981")
NEGATIVE = "#EF4444"            # warning/negative red: only for genuinely negative values

# Chromatic hues a source header fill can be matched to, with the light tint + border used for chips/bands
# (the first four tint/border pairs are the "verified set" in styling.md; the rest are 12% blends).
HUES = {
    "blue": ("#2563EB", "#F0F7FF", "#BAD9F8"),
    "sky": ("#0EA5E9", "#E6F6FA", "#A5DDED"),
    "teal": ("#14B8A6", "#E6F7F5", "#A6DDD6"),
    "green": ("#10B981", "#E4F7EC", "#A4DFC0"),
    "violet": ("#8B5CF6", "#F0EBFA", "#CFC0EE"),
    "amber": ("#F59E0B", "#FEF3E0", "#F8D9A0"),
    "red": ("#EF4444", "#FDECEC", "#F7B9B9"),
}


# Where each palette hue sits on the colour wheel for matching a source fill (any green, yellow-greens included,
# lands on the palette green; the swatch hexes above are not at the centre of their families).
HUE_CENTER = {"blue": 220, "sky": 200, "teal": 175, "green": 125, "violet": 265, "amber": 40, "red": 0}


def resolve(role):
    """Role name (or an already-resolved hex) -> hex; never a role name."""
    if isinstance(role, str) and role.startswith("#"):
        return role
    if role not in ROLES:
        raise ValueError(f"unknown role {role!r}")
    return ROLES[role]


def _hue_of(hex_color):
    r, g, b = (int(hex_color[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, lightness, saturation = colorsys.rgb_to_hls(r, g, b)
    return h * 360, saturation, lightness


def accent_from_fill(fill_hex):
    """Nearest palette hue to a source header fill (``#DDE7F5`` -> blue), so the page feels like the file.

    A grey / white / missing fill has no hue to match and keeps the default ``primary`` blue.
    """
    name = "blue"
    if fill_hex and isinstance(fill_hex, str) and len(fill_hex) == 7:
        hue, saturation, _ = _hue_of(fill_hex)
        if saturation >= 0.08:
            def distance(candidate):
                gap = abs(hue - HUE_CENTER[candidate])
                return min(gap, 360 - gap)
            name = min(HUES, key=distance)
    main, tint, border = HUES[name]
    return {"source": fill_hex, "hue": name, "hex": main, "tint": tint, "border": border}


def categorical_scheme(accent_hex):
    """Positional chart / donut palette: the accent first, the vetted mid hues next, red last (it is only for
    genuinely negative values, so it only appears once the other hues are used up)."""
    rest = [h for h in (ROLES["series-mid"], CATEGORICAL_MID[0], CATEGORICAL_MID[2], CATEGORICAL_MID[4],
                        CATEGORICAL_MID[1], ROLES["series-grey"], CATEGORICAL_MID[3])
            if h.lower() != accent_hex.lower()]
    return [accent_hex] + rest


def settings_theme(accent_hex):
    """Workbook-level theme patch (the only spec path for multi-series and donut/pie slice colors)."""
    if not SURFACES["categorical_scheme"]:
        return {}
    return {"theme": {"overrides": {"categoricalScheme": categorical_scheme(accent_hex)}}}


# ---------------------------------------------------------------- number-format roles (styling.rb format_for)
D3_FORMATS = {"currency": "$,.0f", "integer": ",.0f", "percent": ".1%", "decimal": ",.2f"}


def format_for(role):
    """d3 number format for a semantic role (never an Excel mask: ``$#,##0`` is rejected by Sigma)."""
    return {"kind": "number", "formatString": D3_FORMATS[role]}


# ---------------------------------------------------------------- surfaces
CARD = {"backgroundColor": ROLES["fill"], "borderColor": ROLES["hairline"], "borderWidth": 1,
        "borderRadius": "round"}


def card_style():
    """Neutral card chrome (fill + hairline). Border fields are never combined with ``padding``."""
    return dict(CARD)


def header_band(element_id, title, subtitle, accent):
    """Recipe 1 on a light canvas: a container tinted with the accent holding one text element (title from the
    source title cell + one quiet subline). A text element has no ``style`` in the OpenAPI, so the band is a
    container (``style {backgroundColor, borderRadius}``) with the text laid out inside it.

    Returns ``(container, text)``; the band id is ``<element_id>Band``.
    """
    body = f'# <span style="color: {ROLES["ink"]}">{_markup(title)}</span>'
    if subtitle:
        body += f'\n<span style="color: {ROLES["muted"]}">{_markup(subtitle)}</span>'
    text = {"id": element_id, "kind": "text", "body": body, "verticalAlign": "center"}
    container = {"id": f"{element_id}Band", "kind": "container"}
    if SURFACES["container_style"]:
        container["style"] = {"backgroundColor": accent["tint"], "borderRadius": "round"}
    return container, text


def section_label(element_id, text):
    """Recipe 3, the quiet uppercase variant (a heading would compete with the header band)."""
    return {"id": element_id, "kind": "text",
            "body": f'<p class="p-small"><span style="color: {ROLES["muted"]}">**{_markup(str(text).upper())}**'
                    "</span></p>"}


def divider(element_id):
    return {"id": element_id, "kind": "divider"}


KPI_LABEL_PX, KPI_VALUE_PX = 14, 28


def kpi_card(element_id, label, source, formula, fmt):
    """Recipe 2 without the wrapper container: the KPI carries its own card style. Every card is a peer: a quiet
    muted 14 px title and a bold 28 px value, no accent on the first."""
    column = {"id": f"{element_id}_v", "formula": formula}
    if fmt:
        column["format"] = fmt
    element = {"id": element_id, "kind": "kpi-chart", "source": source, "columns": [column],
               "value": {"columnId": column["id"], "fontSize": KPI_VALUE_PX, "fontWeight": "bold"}}
    element["name"] = ({"text": label, "color": ROLES["muted"], "fontSize": KPI_LABEL_PX,
                        "fontWeight": "normal"} if SURFACES["kpi_name_object"] else label)
    if SURFACES["kpi_card_style"]:
        element["style"] = card_style()
    return element


def kpi_text_card(element_id, label, formula, fmt):
    """A KPI whose formula spans several source elements: a card container holding a text element that
    interpolates the formula (``{{ast | d3 format}}``), laid out like the kpi-chart cards: a muted 14 px title,
    a blank line, a bold 28 px value. Span styles may only use color, background-color, font-size and
    font-family (Sigma rejects font-weight), so the weight is markdown bold.
    Returns ``(container, text)``; the container carries the id of the KPI."""
    spec = (fmt or {}).get("formatString")
    value = f"{{{{{formula} | {spec}}}}}" if spec else f"{{{{{formula}}}}}"
    body = (f'<span style="color: {ROLES["muted"]}; font-size: {KPI_LABEL_PX}px">{_markup(label)}</span>\n\n'
            f'<span style="font-size: {KPI_VALUE_PX}px">**{value}**</span>')
    text = {"id": f"{element_id}_t", "kind": "text", "verticalAlign": "center", "body": body}
    container = {"id": element_id, "kind": "container", "style": card_style()}
    return container, text


SPAN_PROPERTIES = {"color", "background-color", "font-size", "font-family"}


def disallowed_span_styles(body):
    """CSS properties in a text body's ``<span style>`` that Sigma does not accept (empty when all are fine)."""
    import re
    bad = set()
    for style in re.findall(r'<span style="([^"]*)"', body or ""):
        for declaration in filter(None, (d.strip() for d in style.split(";"))):
            prop = declaration.split(":", 1)[0].strip().lower()
            if prop not in SPAN_PROPERTIES:
                bad.add(prop)
    return sorted(bad)


def table_style(header, kind="table"):
    """``tableStyle.textStyles`` header fill / weight / color from the source header cells.

    ``header`` is ``{"fill": "#DDE7F5", "bold": True, "color": None}`` (any key optional). A pivot also
    gets ``columnHeader`` (its column-header band); a flat table uses ``header``.
    """
    if not header or not SURFACES["table_text_styles"]:
        return None
    style = {}
    if header.get("fill"):
        style["backgroundColor"] = header["fill"]
    if header.get("color"):
        style["color"] = header["color"]
    if header.get("bold"):
        style["fontWeight"] = "bold"
    if not style:
        return None
    text_styles = {"header": dict(style)}
    if kind == "pivot-table":
        text_styles["columnHeader"] = dict(style)
    return {"textStyles": text_styles}


def negative_red_rule(column_ids):
    """Excel's ``[Red]`` negative section: a single-condition rule coloring the cell text when < 0."""
    if not column_ids or not SURFACES["conditional_single"]:
        return None
    return {"type": "single", "columnIds": list(column_ids), "condition": "<", "value": 0,
            "style": {"color": NEGATIVE}}


def chart_card():
    return card_style() if SURFACES["chart_card_style"] else None


def single_color(hex_color):
    return {"by": "single", "value": hex_color} if SURFACES["chart_color_single"] else None


def _markup(text):
    """Body text for a span: markdown/HTML specials neutralised (titles come from workbook cells)."""
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace("*", "\\*").replace("{", "(").replace("}", ")")
            .replace("\n", " ").strip())


# ---------------------------------------------------------------- grid geometry (24-column page)
PAGE_COLS = 24
HEADER_ROWS = 4
KPI_ROWS = 7
CHART_ROWS = 16
LABEL_ROWS = 2
DIVIDER_ROWS = 1
GRID_ROW_PX = 23           # measured on a live render: 20 grid rows ~ 464 px
ROW_PX = 20                # one data row of the spreadsheet-preset table
CHROME_PX = 140            # title + header + summary bar + card padding of a table
MIN_TABLE_ROWS, MAX_TABLE_ROWS = 9, 40


def table_height(n_rows, extra_header_rows=0):
    """Grid rows for a table / pivot that shows ``n_rows`` data rows (sized to its rows, one row of slack)."""
    pixels = CHROME_PX + ROW_PX * (max(n_rows, 1) + extra_header_rows)
    rows = math.ceil(pixels / GRID_ROW_PX) + 1
    return max(MIN_TABLE_ROWS, min(MAX_TABLE_ROWS, rows))
