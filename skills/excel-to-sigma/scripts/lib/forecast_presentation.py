"""Presentation of a forecast plan: source number formats, title, chart de-duplication, page composition.

Everything here is presentation only. No formula, parity target, KPI value or section sum is touched: the drafter
reads the source cells' number formats / header looks / title and the builder styles the same elements it always
emitted. The shared styling lives in ``presentation_style`` (the table-report builder uses the same recipes).

Plan field ``presentation`` (absent on a plan drafted before this module: the builder then emits the legacy page)::

    {"version": 1, "title", "subtitle", "accent": {...}, "look"?: <reference look>,
     "headers": {"pivot"|"gl"|"assumptions": {fill, bold, color}},
     "formats": {"money": E, "percent": E, "sections": {name: E}, "derived": {label: E}, "lines": {line: E},
                 "gl": E, "glDate": {datetime format}, "drivers": E},
     "notes": [str]}
    E = {"code": <Excel code>, "cell": <source cell>, "format"?: <Sigma format>, "negativeRed"?: true}

An entry without ``format`` is a General / unmapped code: the builder falls back to the ``--currency-format`` /
``--pct-format`` plan fields (the CLI defaults), never to a format of its own.
"""
from __future__ import annotations

import re
from collections import Counter

try:
    from . import excel_formats as ef
    from . import presentation_style as ps
    from . import reference_look as rl
except ImportError:  # direct import when scripts/lib itself is on sys.path
    import excel_formats as ef
    import presentation_style as ps
    import reference_look as rl

VERSION = 1
PIVOT_HEADER_ROWS = 3          # Period / Month group row, Actual-Forecast row, period row of the pivot
INPUT_ROWS = (16, 20)          # rate input, manual input (an input table is not a presentation table)
HEX6 = re.compile(r"[0-9A-Fa-f]{6}")


# ---------------------------------------------------------------- drafter side: read the source cells
def code_of(cell):
    return str(cell.number_format) if getattr(cell, "number_format", None) else "General"


def _dominant(cells):
    """``(code, address)`` of the most common non-General format among ``cells`` that hold a value; General when
    every cell is General (or empty)."""
    counts, first = Counter(), {}
    for cell in cells:
        if cell.value is None:
            continue
        code = code_of(cell)
        if code.strip().lower() == "general":
            continue
        counts[code] += 1
        first.setdefault(code, f"{cell.parent.title}!{cell.coordinate}")
    if not counts:
        return "General", None
    code = counts.most_common(1)[0][0]
    return code, first[code]


class _Notes:
    """Excel format codes that did not translate exactly (``NUMBER_FORMAT_APPROXIMATED``)."""

    def __init__(self):
        self.codes = {}

    def add(self, code, reasons, where):
        if not reasons:
            return
        note = self.codes.setdefault(str(code), {"reasons": [], "where": []})
        note["reasons"] += [r for r in reasons if r not in note["reasons"]]
        if where and where not in note["where"]:
            note["where"].append(where)

    def evidence(self):
        return [f"{code}: {'; '.join(note['reasons'])} (at {', '.join(note['where'][:3])})"
                for code, note in self.codes.items()]


def _entry(code, where, notes):
    """Plan format entry for source cells that use ``code`` (see the module docstring)."""
    entry = {"code": code}
    if where:
        entry["cell"] = where
    result = ef.translate(code)
    if result.format and not result.is_date:
        entry["format"] = result.format
        if result.negative_red:
            entry["negativeRed"] = True
        notes.add(code, result.approximated, where)
    elif result.category not in ("general", "text"):
        notes.add(code, result.approximated or ["unsupported code"], where)
    return entry


def _row_cells(ws, row, cols):
    return [ws.cell(row, col) for col in cols]


def _title(ws, header_row):
    """The report sheet's title cell: A1, else the first text cell of column A above the header row."""
    for row in range(1, max(header_row, 2)):
        value = ws.cell(row, 1).value
        if isinstance(value, str) and value.strip() and not value.startswith("="):
            return re.sub(r"\s+", " ", value).strip()
    return None


def _gl_cells(gl, header_row, col, limit=400):
    return [gl.cell(r, col) for r in range(header_row + 1, min(gl.max_row, header_row + limit) + 1)]


def propose(*, src, wf, ws, header_row, period_cols, line_rows, totals, derived, gl, gl_header, gl_cols,
            driver, look=None):
    """``(presentation, notes)``: the plan's ``presentation`` block and the :class:`_Notes` of format codes that
    could not be translated exactly. ``driver`` is ``(sheet, first row, rate column)`` of the driver table or
    ``None``; ``line_rows`` / ``totals`` / ``derived`` are the drafter's own row tables."""
    notes = _Notes()
    formats = {}
    # one value column of the pivot: the dominant format over the line and section-total rows
    amount_rows = [row for row, _, _ in line_rows] + [row for row, _ in totals.values()]
    code, where = _dominant([c for row in amount_rows for c in _row_cells(ws, row, period_cols)])
    formats["money"] = _entry(code, where, notes)
    formats["sections"] = {}
    for section, (row, _) in totals.items():
        code, where = _dominant(_row_cells(ws, row, period_cols))
        formats["sections"][section] = _entry(code, where, notes)
    formats["derived"] = {}
    for d in derived:
        code, where = _dominant(_row_cells(ws, d["sourceRow"], period_cols))
        formats["derived"][d["label"]] = _entry(code, where, notes)
    percents = Counter()
    for d in derived:
        if d["kind"] == "ratio":
            entry = formats["derived"][d["label"]]
            if ef.is_percent_format(entry.get("format")):
                percents[entry["code"]] += 1
    if percents:
        best = percents.most_common(1)[0][0]
        formats["percent"] = next(e for d in derived if d["kind"] == "ratio"
                                  for e in [formats["derived"][d["label"]]] if e["code"] == best)
    # a line whose own format differs from the pivot's value format is listed (a pivot has one value format)
    formats["lines"] = {}
    mixed = []
    for row, label, _ in line_rows:
        line_code, line_where = _dominant(_row_cells(ws, row, period_cols))
        if line_code != formats["money"]["code"] and line_code.strip().lower() != "general":
            formats["lines"][label] = _entry(line_code, line_where, notes)
            mixed.append(f"{label}: {line_code}")
    if mixed:
        notes.add(formats["money"]["code"],
                  [f"the P&L pivot has one value format, so {len(mixed)} line(s) with another format show this "
                   f"one ({'; '.join(mixed[:4])}); values are unchanged"], f"{ws.title}!lines")
    if gl is not None and gl_cols.get("amt"):
        code, where = _dominant(_gl_cells(gl, gl_header, gl_cols["amt"]))
        formats["gl"] = _entry(code, where, notes)
        if gl_cols.get("date"):
            code, _ = _dominant(_gl_cells(gl, gl_header, gl_cols["date"]))
            result = ef.translate(code)
            if result.is_date and result.format:
                formats["glDate"] = result.format
                notes.add(code, result.approximated, f"{gl.title}!date")
    if driver:
        dsheet, first_row, rate_col = driver
        dws = wf[dsheet]
        cells = [dws.cell(r, rate_col) for r in range(first_row, dws.max_row + 1)
                 if isinstance(dws.cell(r, rate_col).value, (int, float))]
        code, where = _dominant(cells)
        formats["drivers"] = _entry(code, where, notes)
    headers = {}
    head = ws.cell(header_row, 1)
    style = ps.cell_header_style(head, keep_light_text=True) or ps.cell_header_style(
        ws.cell(header_row, period_cols[0]), keep_light_text=True)
    if style:
        headers["pivot"] = style
    if gl is not None and gl_cols:
        style = ps.cell_header_style(gl.cell(gl_header, min(gl_cols.values())), keep_light_text=True)
        if style:
            headers["gl"] = style
    if driver:
        style = ps.cell_header_style(wf[driver[0]].cell(driver[1] - 1, 1), keep_light_text=True) if driver[1] > 1 else None
        if style:
            headers["assumptions"] = style
    accent = ps.accent_from_fill((headers.get("pivot") or {}).get("fill"))
    modified = getattr(wf.properties, "modified", None)
    parts = [src.name, f"{ws.title} sheet"] + ([f"as of {modified.date().isoformat()}"] if modified else [])
    presentation = {"version": VERSION, "title": _title(ws, header_row), "subtitle": " · ".join(parts),
                    "accent": accent, "headers": headers, "formats": formats}
    ratio_rows = [d["label"] for d in derived if d["kind"] == "ratio"]
    presentation["notes"] = (
        [f"Source ratio row(s) {', '.join(ratio_rows)} are not rows of the P&L pivot (it lists lines and section "
         "totals); the ratio KPI card carries the headline ratio."] if ratio_rows else [])
    if look is not None:
        presentation["look"] = look
    return presentation, notes


def series_colors(chart):
    """Per-series source colours (``#RRGGBB``) of an openpyxl chart, or ``None`` when no series sets one."""
    found = []
    for series in chart.series:
        props = series.graphicalProperties
        fill = None
        if props is not None:
            for candidate in (props.line.solidFill if props.line is not None else None, props.solidFill):
                value = getattr(candidate, "srgbClr", None)
                value = getattr(value, "val", value)
                if isinstance(value, str) and HEX6.fullmatch(value):
                    fill = "#" + value.upper()
                    break
        found.append(fill)
    return found if any(found) else None


# ---------------------------------------------------------------- builder side: formats of a plan
class Formats:
    """The Sigma number formats a plan's elements take. A plan without ``presentation`` (drafted before it) gets
    ``cur`` / ``pct`` (the CLI defaults) everywhere, i.e. exactly the legacy emission."""

    def __init__(self, plan, cur, pct):
        pres = plan.get("presentation") or {}
        self.pres = pres
        self.cur, self.pct = cur, pct
        self.formats = pres.get("formats") or {}

    @staticmethod
    def _use(entry, fallback):
        return dict(entry["format"]) if entry and entry.get("format") else dict(fallback)

    @property
    def money(self):
        return self._use(self.formats.get("money"), self.cur)

    def section(self, name):
        return self._use((self.formats.get("sections") or {}).get(name), self.money)

    def derived(self, label):
        return self._use((self.formats.get("derived") or {}).get(label), self.money)

    def ratio(self, label):
        """The Gross-Margin-% style card: the source row of that name if it shows a percentage, else the plan's
        percentage format, else ``--pct-format``."""
        entry = (self.formats.get("derived") or {}).get(label)
        if entry and ef.is_percent_format(entry.get("format")):
            return dict(entry["format"])
        entry = self.formats.get("percent")
        return dict(entry["format"]) if entry and ef.is_percent_format(entry.get("format")) else dict(self.pct)

    @property
    def gl(self):
        """GL amounts: their own source format; a General GL column shows like the P&L cells it feeds (the
        source's own format of those sums), then ``--currency-format``."""
        return self._use(self.formats.get("gl"), self.money)

    @property
    def gl_date(self):
        return dict(self.formats["glDate"]) if self.formats.get("glDate") else None

    @property
    def drivers(self):
        return self._use(self.formats.get("drivers"), self.pct)

    def red(self, key):
        return bool((self.formats.get(key) or {}).get("negativeRed"))


# ---------------------------------------------------------------- charts
def trend_covered(plan, trend_keys, grain):
    """The title of a mapped source chart that already shows the generated trend (the same series at the plan's
    period grain over every period), else ``None``. ``trend_keys`` are the trend's series named like
    ``sourceCharts[].series`` (the section total label / amount row label)."""
    periods = plan.get("periods") or []
    if plan.get("presentation") is None:
        return None
    wanted = set(trend_keys)
    for chart in plan.get("sourceCharts") or []:
        if chart.get("grain") != ("month" if grain == "month" else "period"):
            continue
        if not wanted <= set(chart.get("series") or []):
            continue
        if chart["grain"] == "month" and periods:
            first, last = periods[0].get("end", ""), periods[-1].get("end", "")
            if not (chart.get("periodStart", "") <= first[:8] + "01" and chart.get("periodEnd", "") >= last):
                continue
        return chart["title"]
    return None


def scheme_from_source(plan):
    """The categorical colour scheme of the first source chart whose every series sets a colour (those colours
    first, then the default palette), else ``None``."""
    for chart in plan.get("sourceCharts") or []:
        colors = chart.get("colors")
        if colors and all(colors):
            rest = [c for c in ps.categorical_scheme(ps.ROLES["primary"]) if c.lower() not in {x.lower() for x in colors}]
            return list(colors) + rest
    return None


# ---------------------------------------------------------------- page composition
def check_look(plan):
    return rl.problem((plan.get("presentation") or {}).get("look"))


def _pivot_style(pivot, plan, fm, look):
    headers = (plan["presentation"].get("headers") or {})
    ids = [c["id"] for c in pivot["columns"] if c["id"] in pivot["values"]] if fm.red("money") else []
    ps.style_table(pivot, headers.get("pivot"), ids, look)
    if pivot.get("conditionalFormats") and ps.SURFACES["pivot_negative_scope"]:
        pivot["conditionalFormats"] = [{**rule, "includeValues": True, "includeSubtotals": True,
                                        "includeGrandTotals": True} for rule in pivot["conditionalFormats"]]
    if ps.SURFACES["pivot_subtotal_bold"]:
        pivot["totals"] = {"subTotalFontWeight": "bold"}       # the section rows (Excel's bold section totals)


def _table_style(table, plan, key, red_ids, look):
    ps.style_table(table, (plan["presentation"].get("headers") or {}).get(key), red_ids, look)


def compose(plan, *, kpis, charts, pivot, gl, assump, inputs, hidden, hidden_ids, fm, n_pivot_rows):
    """The workbook pages and layout: header band, KPI row, hero chart(s), divider + DETAIL label, P&L pivot; the
    GL detail page; the drivers / editable-input page (omitted when it would hold nothing); the hidden model.

    Returns ``{"pages": [...], "layout": <xml>, "settings": {...}}``. ``kpis`` / ``charts`` are built elements in
    display order; the layout covers every element once."""
    pres = plan["presentation"]
    look = pres.get("look")
    accent = pres.get("accent") or ps.accent_from_fill(None)
    roomy = ps.roomy_tables(look)
    band, header = ps.header_band("pnlHeader", pres.get("title") or plan["name"], pres.get("subtitle"), accent, look)
    main, placements = [band, header], []
    placements.append((band["id"], 1, ps.PAGE_COLS + 1, 1, 1 + ps.HEADER_ROWS,
                       [(header["id"], 1, ps.PAGE_COLS + 1, 1, 1 + ps.HEADER_ROWS)]))
    cursor = 1 + ps.HEADER_ROWS
    if kpis:
        width = ps.PAGE_COLS // len(kpis)
        for number, card in enumerate(kpis):
            placements.append((card["id"], 1 + number * width, 1 + (number + 1) * width, cursor,
                               cursor + ps.KPI_ROWS))
        main += kpis
        cursor += ps.KPI_ROWS
    main += charts
    if len(charts) == 1:
        placements.append((charts[0]["id"], 1, ps.PAGE_COLS + 1, cursor, cursor + ps.CHART_ROWS))
        cursor += ps.CHART_ROWS
    elif len(charts) == 2:             # both exist after de-duplication: side by side
        for chart, (c0, c1) in zip(charts, ((1, 13), (13, 25))):
            placements.append((chart["id"], c0, c1, cursor, cursor + ps.CHART_ROWS))
        cursor += ps.CHART_ROWS
    elif charts:                        # a full-width hero, then the others in pairs
        placements.append((charts[0]["id"], 1, ps.PAGE_COLS + 1, cursor, cursor + ps.CHART_ROWS))
        cursor += ps.CHART_ROWS
        for index in range(1, len(charts), 2):
            row = charts[index:index + 2]
            splits = [(1, 25)] if len(row) == 1 else [(1, 13), (13, 25)]
            for chart, (c0, c1) in zip(row, splits):
                placements.append((chart["id"], c0, c1, cursor, cursor + ps.CHART_ROWS))
            cursor += ps.CHART_ROWS
    if kpis or charts:
        divider, label = ps.divider("pnlDivider"), ps.section_label("pnlDetail", "Detail", look)
        placements.append((divider["id"], 1, ps.PAGE_COLS + 1, cursor, cursor + ps.DIVIDER_ROWS))
        cursor += ps.DIVIDER_ROWS
        placements.append((label["id"], 1, ps.PAGE_COLS + 1, cursor, cursor + ps.LABEL_ROWS))
        cursor += ps.LABEL_ROWS
        main += [divider, label]
    _pivot_style(pivot, plan, fm, look)
    height = ps.table_height(n_pivot_rows, extra_header_rows=PIVOT_HEADER_ROWS, roomy=roomy)
    placements.append((pivot["id"], 1, ps.PAGE_COLS + 1, cursor, cursor + height))
    main.append(pivot)

    gl_ids = [c["id"] for c in gl["columns"] if c["name"] == "Amount"] if fm.red("gl") else []
    _table_style(gl, plan, "gl", gl_ids, look)
    detail_height = ps.table_height(len(plan.get("actuals") or []), roomy=roomy)
    detail = [(gl["id"], 1, ps.PAGE_COLS + 1, 1, 1 + detail_height)]

    input_elements = ([assump] if assump is not None else []) + list(inputs)
    pages = [{"id": "pnl", "name": "P&L Overview", "elements": main},
             {"id": "detail", "name": "Detail", "elements": [gl]}]
    layout_pages = [("pnl", placements), ("detail", detail)]
    if input_elements:                 # an Inputs page that would hold nothing is omitted
        spans, row = [], 1
        if assump is not None:
            _table_style(assump, plan, "assumptions", [], look)
            rows = ps.table_height(len(plan.get("assumptions") or []), roomy=roomy)
            spans.append((assump["id"], 1, ps.PAGE_COLS + 1, row, row + rows))
            row += rows
        for element, rows in zip(inputs, INPUT_ROWS):
            spans.append((element["id"], 1, ps.PAGE_COLS + 1, row, row + rows))
            row += rows
        pages.append({"id": "inputs", "name": "Inputs", "elements": input_elements})
        layout_pages.append(("inputs", spans))
    pages.append({"id": "model", "name": "Model (sources)", "visibility": "hidden", "elements": hidden})
    model = [(element_id, 1, ps.PAGE_COLS + 1, 1 + 12 * n, 13 + 12 * n) for n, element_id in enumerate(hidden_ids)]
    layout_pages.append(("model", model))
    charts_present = bool(charts)
    settings = ps.settings_theme(accent["hex"], look) if (charts_present or look is not None) else {}
    scheme = scheme_from_source(plan) if look is None and charts_present else None
    if scheme:
        settings = {"theme": {"overrides": {"categoricalScheme": scheme}}}
    return {"pages": pages, "layout": ps.layout_xml(layout_pages), "settings": settings}
